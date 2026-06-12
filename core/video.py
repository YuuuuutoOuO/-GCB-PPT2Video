"""
core/video.py
─────────────
影片合成：截圖匯出、單頁 MP4 合成、ffmpeg 串接
"""

import os
import sys
import subprocess

try:
    from moviepy.editor import ImageClip, AudioFileClip   # moviepy 1.x
except ImportError:
    from moviepy import ImageClip, AudioFileClip           # moviepy 2.x

from core.utils import extract_title_number, set_pptx_export_dpi


# ══════════════════════════════════════════════════════════════
# 投影片截圖（COM 介面，Windows only）
# ══════════════════════════════════════════════════════════════

def export_slides_to_images(pptx_path: str, image_dir: str, dpi: int = 200):
    """
    透過 PowerPoint COM 將每頁匯出為 JPG。
    匯出前先透過 Registry 設定 DPI 以提升畫質。
    """
    # 設定匯出 DPI
    set_pptx_export_dpi(dpi)

    import win32com.client
    pptx_abs = os.path.abspath(pptx_path)
    img_abs  = os.path.abspath(image_dir)

    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(pptx_abs)
    deck.SaveAs(img_abs, 17)   # 17 = ppSaveAsJPG
    deck.Close()
    powerpoint.Quit()


def get_sorted_images(image_dir: str) -> list:
    """語系無關的純數字排序，相容中英文檔名"""
    files = [f for f in os.listdir(image_dir) if f.lower().endswith(".jpg")]
    return sorted(files, key=lambda x: int("".join(filter(str.isdigit, x)) or 0))


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

def create_single_clip(
    img_path: str,
    audio_path: str,
    out_path: str,
    fps: int,
) -> bool:
    """
    合成單頁 MP4，完成後釋放資源。
    回傳 True 代表成功，False 代表失敗。
    """
    # 先確認音檔存在且大小正常
    if not os.path.exists(audio_path) or os.path.getsize(audio_path) < 1024:
        print(f"\n  [錯誤] 音檔異常，跳過：{os.path.basename(audio_path)}")
        return False

    try:
        audio_clip = AudioFileClip(audio_path)
        image_clip = (
            ImageClip(img_path)
            .set_duration(audio_clip.duration)
            .set_audio(audio_clip)
        )
        image_clip.write_videofile(
            out_path, fps=fps,
            codec="libx264", audio_codec="aac",
            logger=None,
        )
        audio_clip.close()
        image_clip.close()
        return True

    except Exception as e:
        print(f"\n  [錯誤] 合成失敗：{e}")
        if os.path.exists(out_path):
            os.remove(out_path)   # 清除寫到一半的殘檔
        return False


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
        print('例如：ffmpeg_path: "C:\\\\ffmpeg\\\\bin\\\\ffmpeg.exe"')
        sys.exit(1)

    os.remove(concat_list_path)
