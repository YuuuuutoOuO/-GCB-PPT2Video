"""
core/video.py
─────────────
影片合成：截圖匯出、單頁 MP4 合成、ffmpeg 串接

優化紀錄：
  - 移除 moviepy 依賴，改用直接呼叫 ffmpeg subprocess
  - 支援 NVIDIA GPU 編碼（h264_nvenc），CPU fallback（libx264）
  - Phase C 改為 ThreadPoolExecutor 平行合成
  - GPU/CPU 使用不同 level：nvenc 最低 3.1，libx264 用 3.0
"""

import os
import sys
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed

from core.utils import extract_title_number, set_pptx_export_dpi


# ══════════════════════════════════════════════════════════════
# 投影片截圖（COM 介面，Windows only）
# ══════════════════════════════════════════════════════════════

def export_slides_to_images(
    pptx_path: str,
    image_dir: str,
    target_pages: list[int],
    dpi: int = 200,
) -> dict[int, str]:
    """
    透過 PowerPoint COM 只截指定頁的截圖（1-based 頁碼）。
    回傳 {0-based-idx: image_path} 的 dict。
    """
    from tqdm import tqdm
    set_pptx_export_dpi(dpi)

    import win32com.client
    pptx_abs = os.path.abspath(pptx_path)
    img_abs  = os.path.abspath(image_dir)

    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(pptx_abs)

    image_map = {}

    pbar = tqdm(
        total=len(target_pages),
        desc="  Phase A 截圖",
        unit="頁",
        ncols=70,
        bar_format="{desc}：{n_fmt}/{total_fmt} {bar} {percentage:3.0f}% [{elapsed}<{remaining}]",
    )

    for page_1based in target_pages:
        img_path = os.path.join(img_abs, f"slide_{page_1based:04d}.jpg")
        if not os.path.exists(img_path):
            slide = deck.Slides(page_1based)
            slide.Export(img_path, "JPG")
        image_map[page_1based - 1] = img_path
        pbar.update(1)

    pbar.close()
    deck.Close()
    powerpoint.Quit()

    return image_map


# ══════════════════════════════════════════════════════════════
# clip 輸出路徑
# ══════════════════════════════════════════════════════════════

def get_clip_path(parsed: dict, page_1based: int, clips_dir: str) -> str:
    """
    依標題數字決定子資料夾：
      有數字 → clips/058/page_XXXX.mp4
      無數字 → clips/intro/page_XXXX.mp4
    """
    num = extract_title_number(parsed["title"])
    folder = os.path.join(clips_dir, f"{num:03d}" if num >= 0 else "intro")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, f"page_{page_1based:04d}.mp4")


# ══════════════════════════════════════════════════════════════
# 單頁 MP4 合成
# ══════════════════════════════════════════════════════════════

def _build_ffmpeg_cmd(
    img_path: str,
    audio_path: str,
    out_path: str,
    fps: int,
    ffmpeg_path: str,
    use_gpu: bool,
) -> list[str]:
    """
    組合 ffmpeg 指令列表。
    GPU/CPU 使用不同的 level：
      GPU (nvenc) : baseline + level 5.0（nvenc 不支援 3.0）
      CPU (x264)  : baseline + level 3.0（相容性最廣）
    """
    codec  = "h264_nvenc" if use_gpu else "libx264"
    preset = "p4"         if use_gpu else "veryfast"
    level  = "5.0"        if use_gpu else "3.0"

    cmd = [
        ffmpeg_path, "-y",
        "-loop", "1", "-i", img_path,
        "-i", audio_path,
        "-c:v", codec,
        "-preset", preset,
        "-vf", "scale=trunc(iw/2)*2:trunc(ih/2)*2",
        "-pix_fmt", "yuv420p",
        "-profile:v", "baseline",
        "-level:v", level,
        "-c:a", "aac",
        "-shortest",
        "-fps_mode", "vfr",
        out_path,
    ]

    # nvenc 額外加 tune hq 提升畫質
    if use_gpu:
        cmd.insert(cmd.index("-c:a"), "-tune")
        cmd.insert(cmd.index("-tune") + 1, "hq")

    return cmd


def create_single_clip(
    img_path: str,
    audio_path: str,
    out_path: str,
    fps: int,
    ffmpeg_path: str = "ffmpeg",
    use_gpu: bool = True,
) -> bool:
    """
    直接呼叫 ffmpeg 合成單頁 MP4。
    GPU 失敗時一律自動 fallback 到 CPU。
    回傳 True 代表成功，False 代表失敗。
    """
    if not os.path.exists(audio_path) or os.path.getsize(audio_path) < 1024:
        print(f"\n  [錯誤] 音檔異常，跳過：{os.path.basename(audio_path)}")
        return False

    cmd = _build_ffmpeg_cmd(img_path, audio_path, out_path, fps, ffmpeg_path, use_gpu)

    try:
        subprocess.run(cmd, check=True, capture_output=True)
        return True

    except subprocess.CalledProcessError as e:
        err_msg = e.stderr.decode("utf-8", errors="ignore")

        # 過濾出真正的錯誤行
        err_lines = [
            line for line in err_msg.splitlines()
            if any(kw in line.lower() for kw in [
                "error", "invalid", "failed", "cannot", "no such",
                "unknown", "not found", "unsupported", "denied",
                "nvenc", "cuda", "gpu",
            ])
        ]
        err_summary = "\n    ".join(err_lines[:5]) if err_lines else err_msg[-300:]

        if use_gpu:
            print(f"\n  [警告] GPU 編碼失敗，自動改用 CPU 重試：{os.path.basename(out_path)}")
            print(f"    原因：{err_summary}")
            if os.path.exists(out_path):
                os.remove(out_path)
            return create_single_clip(
                img_path, audio_path, out_path,
                fps, ffmpeg_path, use_gpu=False,
            )

        print(f"\n  [錯誤] CPU 合成也失敗：{os.path.basename(out_path)}")
        print(f"    原因：{err_summary}")
        if os.path.exists(out_path):
            os.remove(out_path)
        return False

    except FileNotFoundError:
        print(f"\n  [錯誤] 找不到 ffmpeg：{ffmpeg_path}")
        return False


# ══════════════════════════════════════════════════════════════
# 平行合成多頁（Phase 1 使用）
# ══════════════════════════════════════════════════════════════

def create_clips_parallel(
    todo: list,
    image_map: dict[int, str],
    audio_map: dict,
    parsed_list: list,
    clips_dir: str,
    fps: int,
    ffmpeg_path: str,
    use_gpu: bool,
    max_workers: int = 4,
) -> tuple[int, int]:
    """
    平行合成多頁 MP4。
    todo      : 待處理的 0-based index 列表
    image_map : {0-based-idx: image_path}
    回傳 (done_count, fail_count)
    """
    from tqdm import tqdm

    done_count = 0
    fail_count = 0

    def _task(idx):
        img_path   = image_map.get(idx, "")
        audio_path = audio_map.get(idx, "")
        clip_path  = get_clip_path(parsed_list[idx], idx + 1, clips_dir)
        return create_single_clip(
            img_path, audio_path, clip_path,
            fps, ffmpeg_path, use_gpu,
        )

    pbar = tqdm(
        total=len(todo),
        desc="  Phase C 合成",
        unit="頁",
        ncols=70,
        bar_format="{desc}：{n_fmt}/{total_fmt} {bar} {percentage:3.0f}% [{elapsed}<{remaining}]",
    )

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_task, idx): idx for idx in todo}
        for future in as_completed(futures):
            idx = futures[future]
            try:
                success = future.result()
            except Exception as e:
                pbar.write(f"  [錯誤] 第 {idx + 1} 頁 exception：{e}")
                success = False

            if success:
                done_count += 1
            else:
                fail_count += 1

            pbar.update(1)

    pbar.close()
    return done_count, fail_count


# ══════════════════════════════════════════════════════════════
# ffmpeg 串接（phase2 使用）
# ══════════════════════════════════════════════════════════════

def compose_video(
    clips: list,
    output_video: str,
    clips_dir: str,
    ffmpeg_path: str = "ffmpeg",
):
    """產生 concat_list.txt 並用 ffmpeg 無損串接所有 clip"""
    concat_list_path = os.path.join(clips_dir, "concat_list.txt")

    with open(concat_list_path, "w", encoding="utf-8") as f:
        for clip_path in clips:
            f.write(f"file '{os.path.abspath(clip_path)}'\n")

    print(f"\n  ffmpeg 串接中（共 {len(clips)} 個 clip）...")
    print(f"  使用 ffmpeg：{ffmpeg_path}")

    try:
        subprocess.run([
            ffmpeg_path, "-y",
            "-f", "concat", "-safe", "0",
            "-i", concat_list_path,
            "-c", "copy",
            output_video,
        ], check=True)
    except FileNotFoundError:
        print(f"\n[錯誤] 找不到 ffmpeg：{ffmpeg_path}")
        print("請確認 settings.yaml 中的 ffmpeg_path 路徑是否正確。")
        sys.exit(1)

    os.remove(concat_list_path)