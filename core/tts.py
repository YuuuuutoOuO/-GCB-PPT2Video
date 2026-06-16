"""
core/tts.py
───────────
TTS 配音：使用 Edge-TTS 將段落列表轉換為音檔
每頁流程：段落逐一生成音檔 → 段落間插入靜音 → ffmpeg 串接成整頁音檔
支援分批異步生成、Semaphore 限流、失敗自動重試、斷點續跑
失敗的段落會寫入 debug/ 資料夾供事後檢查
"""

import os
import asyncio
import subprocess
from datetime import datetime
import edge_tts


# ══════════════════════════════════════════════════════════════
# Debug 記錄
# ══════════════════════════════════════════════════════════════

_debug_log_path: str = ""   # 由 generate_audio_for_targets 初始化


def _init_debug_log(base_dir: str) -> str:
    """
    建立 debug 資料夾與本次執行的 log 檔。
    檔名含時間戳，避免每次覆蓋。
    回傳 log 檔路徑。
    """
    debug_dir = os.path.join(base_dir, "debug")
    os.makedirs(debug_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path  = os.path.join(debug_dir, f"tts_errors_{timestamp}.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"TTS 錯誤記錄 — {timestamp}\n")
        f.write("=" * 60 + "\n\n")
    return log_path


def _write_debug(log_path: str, message: str):
    """將錯誤訊息附加寫入 log 檔（thread-safe append）"""
    if not log_path:
        return
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(message + "\n")
    except Exception:
        pass   # log 失敗不影響主流程


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
# 單段落音檔生成（含 Semaphore 限流、重試與驗證）
# ══════════════════════════════════════════════════════════════

async def _generate_one(
    text: str,
    path: str,
    voice: str,
    sem: asyncio.Semaphore,
    log_path: str,
    page_1based: int,
    seg_idx: int,
    retries: int = 3,
):
    """
    生成單段落音檔（純文字輸入）。
    - sem      ：全域 Semaphore，控制同時進行的 TTS 請求數
    - log_path ：debug log 檔路徑
    - 自動重試最多 retries 次，每次失敗後等待時間遞增
    - 驗證檔案大小 > 1KB，過小視為損毀並重試
    """
    if not isinstance(text, str) or not text.strip():
        msg = (
            f"[型別錯誤] 第 {page_1based} 頁 段落 {seg_idx}\n"
            f"  檔案：{os.path.basename(path)}\n"
            f"  原因：段落文字無效（非 str 或空白）\n"
            f"  內容：{text!r}\n"
        )
        print(f"\n  [錯誤] {msg.splitlines()[0]}")
        _write_debug(log_path, msg)
        return

    last_error = ""
    for attempt in range(retries):
        try:
            async with sem:
                communicate = edge_tts.Communicate(text, voice)
                await communicate.save(path)

            if os.path.exists(path) and os.path.getsize(path) > 1024:
                return  # 成功

            last_error = "生成檔案過小（可能為靜音或損毀）"
            print(f"\n  [警告] 第 {page_1based} 頁 段落 {seg_idx} 第 {attempt + 1} 次{last_error}，重試中...")
            if os.path.exists(path):
                os.remove(path)

        except Exception as e:
            last_error = str(e)
            print(f"\n  [警告] 第 {page_1based} 頁 段落 {seg_idx} TTS 失敗（第 {attempt + 1} 次）：{last_error}，重試中...")

        # 失敗後等待時間遞增：2s → 4s → 6s
        await asyncio.sleep(2 * (attempt + 1))

    # 全部重試失敗 → 寫入 debug log
    msg = (
        f"[TTS 失敗] 第 {page_1based} 頁 段落 {seg_idx}\n"
        f"  檔案：{os.path.basename(path)}\n"
        f"  錯誤：{last_error}\n"
        f"  內容（前200字）：{text[:200]!r}\n"
    )
    print(f"\n  [錯誤] 第 {page_1based} 頁 段落 {seg_idx} 已達最大重試次數，略過（詳見 debug/）")
    _write_debug(log_path, msg)


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
    sem: asyncio.Semaphore,
    log_path: str,
):
    """
    將一頁的段落列表生成為整頁音檔：
      1. 每個段落分別 TTS 生成音檔（受 sem 限流）
      2. 段落間插入靜音
      3. ffmpeg 串接成整頁音檔
    空頁（parts 為空）→ 直接複製靜音當作該頁音檔
    """
    import shutil
    page_1based = page_idx + 1

    # 空頁：複製靜音音檔
    if not parts:
        shutil.copy(silence_path, page_audio_path)
        return

    # 只有一段：直接生成，不需串接
    if len(parts) == 1:
        await _generate_one(parts[0], page_audio_path, voice, sem, log_path, page_1based, 0)
        return

    # 多段：逐段生成 → 插靜音 → 串接
    seg_files = []
    tasks     = []

    for seg_idx, text in enumerate(parts):
        seg_path = os.path.join(seg_dir, f"page_{page_1based:04d}_seg_{seg_idx:03d}.mp3")
        seg_files.append(seg_path)
        if not os.path.exists(seg_path):
            tasks.append(_generate_one(
                text, seg_path, voice, sem, log_path, page_1based, seg_idx
            ))

    if tasks:
        await asyncio.gather(*tasks)

    # 交錯插入靜音：[seg0, silence, seg1, silence, seg2]
    seg_paths = []
    for i, seg_path in enumerate(seg_files):
        if os.path.exists(seg_path):
            seg_paths.append(seg_path)
        if i < len(seg_files) - 1:
            seg_paths.append(silence_path)

    if len(seg_paths) == 1:
        shutil.copy(seg_paths[0], page_audio_path)
    elif seg_paths:
        _concat_audio(seg_paths, page_audio_path, ffmpeg_path)
    else:
        shutil.copy(silence_path, page_audio_path)


# ══════════════════════════════════════════════════════════════
# 批次生成
# ══════════════════════════════════════════════════════════════

async def generate_audio_for_targets(
    scripts: list[list[str]],
    target_indices: set,
    audio_dir: str,
    settings: dict,
    clips_dir: str = "",
) -> dict:
    """
    只為目標頁生成音檔，支援斷點續跑（已存在的整頁音檔跳過）。
    失敗的段落會寫入 debug/ 資料夾的 log 檔。

    scripts    : list[list[str]]，每頁為段落列表
    回傳 {0-based-idx: audio_path} 的 dict。
    """
    from tqdm import tqdm

    voice          = settings.get("tts_voice", "zh-TW-HsiaoChenNeural")
    batch_size     = settings.get("tts_batch_size", 20)
    pause_ms       = settings.get("pause_between_paragraphs_ms", 300)
    ffmpeg_path    = settings.get("ffmpeg_path", "ffmpeg")
    max_concurrent = settings.get("tts_max_concurrent", 5)

    # 建立 debug log（存到 clips/debug/，不放暫存資料夾）
    debug_base = clips_dir if clips_dir else os.path.dirname(audio_dir)
    log_path = _init_debug_log(debug_base)

    # 建立段落暫存資料夾與靜音檔
    seg_dir      = os.path.join(audio_dir, "_segments")
    os.makedirs(seg_dir, exist_ok=True)
    silence_path = os.path.join(seg_dir, f"silence_{pause_ms}ms.mp3")
    _generate_silence(silence_path, pause_ms, ffmpeg_path)

    # 全域 Semaphore：限制同時進行的 TTS 請求數
    sem = asyncio.Semaphore(max_concurrent)

    targets = sorted(target_indices) if target_indices else list(range(len(scripts)))

    # 先建立 audio_map，並找出哪些頁需要生成
    audio_map    = {}
    todo_indices = []
    for idx in targets:
        audio_path = os.path.join(audio_dir, f"slide_{idx + 1:04d}.mp3")
        audio_map[idx] = audio_path
        if not os.path.exists(audio_path):
            todo_indices.append(idx)

    skipped = len(targets) - len(todo_indices)

    pbar = tqdm(
        total=len(targets),
        initial=skipped,
        desc="  Phase B TTS",
        unit="頁",
        ncols=70,
        bar_format="{desc}：{n_fmt}/{total_fmt} {bar} {percentage:3.0f}% [{elapsed}<{remaining}]",
    )

    async def _generate_and_update(idx):
        await _generate_page_audio(
            parts=scripts[idx],
            page_audio_path=audio_map[idx],
            page_idx=idx,
            voice=voice,
            silence_path=silence_path,
            seg_dir=seg_dir,
            ffmpeg_path=ffmpeg_path,
            sem=sem,
            log_path=log_path,
        )
        pbar.update(1)

    for batch_start in range(0, len(todo_indices), batch_size):
        batch = todo_indices[batch_start: batch_start + batch_size]
        await asyncio.gather(*[_generate_and_update(idx) for idx in batch])

    pbar.close()

    # 若 log 檔只有標題（沒有任何錯誤），刪掉避免產生空檔
    with open(log_path, "r", encoding="utf-8") as f:
        content = f.read()
    if content.count("\n") <= 3:
        os.remove(log_path)
    else:
        print(f"\n  ⚠️  有失敗段落，詳見：{log_path}")

    return audio_map