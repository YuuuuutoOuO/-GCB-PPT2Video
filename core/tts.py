"""
core/tts.py
───────────
TTS 配音：使用 Edge-TTS 將段落列表轉換為音檔
每頁流程：段落逐一生成音檔 → 段落間插入靜音 → ffmpeg 串接成整頁音檔
支援分批異步生成、失敗自動重試、斷點續跑
"""

import os
import asyncio
import subprocess
import edge_tts


# ══════════════════════════════════════════════════════════════
# 靜音音檔生成
# ══════════════════════════════════════════════════════════════

def _generate_silence(path: str, duration_ms: int, ffmpeg_path: str):
    """
    用 ffmpeg 生成指定長度的靜音 mp3。
    只在不存在時生成，可重複使用。
    """
    if os.path.exists(path):
        return
    duration_sec = duration_ms / 1000.0
    subprocess.run([
        ffmpeg_path, "-y",
        "-f", "lavfi", "-i", "anullsrc=r=24000:cl=mono",
        "-t", str(duration_sec),
        "-q:a", "9",
        "-acodec", "libmp3lame",
        path,
    ], check=True, capture_output=True)


# ══════════════════════════════════════════════════════════════
# 單段落音檔生成（含重試與驗證）
# ══════════════════════════════════════════════════════════════

async def _generate_one(
    text: str,
    path: str,
    voice: str,
    retries: int = 3,
):
    """
    生成單段落音檔（純文字輸入）。
    - 自動重試最多 retries 次
    - 驗證檔案大小 > 1KB，過小視為損毀並重試
    """
    if not isinstance(text, str) or not text.strip():
        print(f"\n  [錯誤] 段落文字無效（非 str 或空白），跳過：{os.path.basename(path)}")
        return
    for attempt in range(retries):
        try:
            communicate = edge_tts.Communicate(text, voice)
            await communicate.save(path)

            if os.path.exists(path) and os.path.getsize(path) > 1024:
                return  # 成功

            print(f"\n  [警告] 第 {attempt + 1} 次生成檔案過小，重試中...")
            if os.path.exists(path):
                os.remove(path)

        except Exception as e:
            print(f"\n  [警告] TTS 失敗（第 {attempt + 1} 次）：{e}，重試中...")

        await asyncio.sleep(2)

    print(f"\n  [錯誤] {os.path.basename(path)} 已達最大重試次數，略過此段落")


# ══════════════════════════════════════════════════════════════
# 單頁音檔串接
# ══════════════════════════════════════════════════════════════

def _concat_audio(
    segment_paths: list[str],
    out_path: str,
    ffmpeg_path: str,
):
    """
    用 ffmpeg 將多個音檔（段落 + 靜音交錯）串接成整頁音檔。
    segment_paths 已依序排列：[段落1, 靜音, 段落2, 靜音, 段落3, ...]
    """
    concat_list = out_path + ".concat.txt"
    with open(concat_list, "w", encoding="utf-8") as f:
        for p in segment_paths:
            f.write(f"file '{os.path.abspath(p)}'\n")

    subprocess.run([
        ffmpeg_path, "-y",
        "-f", "concat", "-safe", "0",
        "-i", concat_list,
        "-c", "copy",
        out_path,
    ], check=True, capture_output=True)

    os.remove(concat_list)


# ══════════════════════════════════════════════════════════════
# 單頁完整音檔生成
# ══════════════════════════════════════════════════════════════

async def _generate_page_audio(
    parts: list[str],
    page_audio_path: str,
    page_idx: int,
    voice: str,
    silence_path: str,
    seg_dir: str,
    ffmpeg_path: str,
):
    """
    將一頁的段落列表生成為整頁音檔：
      1. 每個段落分別 TTS 生成音檔
      2. 段落間插入靜音
      3. ffmpeg 串接成整頁音檔
    空頁（parts 為空）→ 直接複製靜音當作該頁音檔
    """
    # 空頁：複製靜音音檔（約 1 秒靜音由靜音檔決定，這裡直接用 pause 靜音）
    if not parts:
        import shutil
        shutil.copy(silence_path, page_audio_path)
        return

    # 只有一段：直接生成，不需串接
    if len(parts) == 1:
        await _generate_one(parts[0], page_audio_path, voice)
        return

    # 多段：逐段生成 → 插靜音 → 串接
    seg_paths = []
    tasks     = []
    seg_files = []

    for seg_idx, text in enumerate(parts):
        seg_path = os.path.join(seg_dir, f"page_{page_idx + 1:04d}_seg_{seg_idx:03d}.mp3")
        seg_files.append(seg_path)
        if not os.path.exists(seg_path):
            tasks.append(_generate_one(text, seg_path, voice))

    if tasks:
        await asyncio.gather(*tasks)

    # 交錯插入靜音：[seg0, silence, seg1, silence, seg2]
    for i, seg_path in enumerate(seg_files):
        if os.path.exists(seg_path):
            seg_paths.append(seg_path)
        if i < len(seg_files) - 1:
            seg_paths.append(silence_path)

    if len(seg_paths) == 1:
        # 只有一段成功生成（其餘失敗）
        import shutil
        shutil.copy(seg_paths[0], page_audio_path)
    elif seg_paths:
        _concat_audio(seg_paths, page_audio_path, ffmpeg_path)
    else:
        # 全段失敗，插入靜音
        import shutil
        shutil.copy(silence_path, page_audio_path)


# ══════════════════════════════════════════════════════════════
# 批次生成
# ══════════════════════════════════════════════════════════════

async def generate_audio_for_targets(
    scripts: list[list[str]],
    target_indices: set,
    audio_dir: str,
    settings: dict,
) -> dict:
    """
    只為目標頁生成音檔，支援斷點續跑（已存在的整頁音檔跳過）。

    scripts    : list[list[str]]，每頁為段落列表
    回傳 {0-based-idx: audio_path} 的 dict。
    """
    voice       = settings.get("tts_voice", "zh-TW-HsiaoChenNeural")
    batch_size  = settings.get("tts_batch_size", 20)
    pause_ms    = settings.get("pause_between_paragraphs_ms", 300)
    ffmpeg_path = settings.get("ffmpeg_path", "ffmpeg")

    # 建立段落暫存資料夾與靜音檔
    seg_dir      = os.path.join(audio_dir, "_segments")
    os.makedirs(seg_dir, exist_ok=True)
    silence_path = os.path.join(seg_dir, f"silence_{pause_ms}ms.mp3")
    _generate_silence(silence_path, pause_ms, ffmpeg_path)

    targets   = sorted(target_indices) if target_indices else list(range(len(scripts)))
    audio_map = {}

    for batch_start in range(0, len(targets), batch_size):
        batch = targets[batch_start: batch_start + batch_size]
        tasks = []

        for idx in batch:
            audio_path = os.path.join(audio_dir, f"slide_{idx + 1:04d}.mp3")
            audio_map[idx] = audio_path

            if not os.path.exists(audio_path):
                tasks.append(_generate_page_audio(
                    parts=scripts[idx],
                    page_audio_path=audio_path,
                    page_idx=idx,
                    voice=voice,
                    silence_path=silence_path,
                    seg_dir=seg_dir,
                    ffmpeg_path=ffmpeg_path,
                ))

        if tasks:
            await asyncio.gather(*tasks)

        done = min(batch_start + batch_size, len(targets))
        print(f"  TTS 進度：{done}/{len(targets)}", end="\r")

    print()
    return audio_map