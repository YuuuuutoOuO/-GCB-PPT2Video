"""
core/tts.py
───────────
TTS 配音：使用 Edge-TTS 將 SSML 文字轉換為音檔
支援分批異步生成、失敗自動重試、斷點續跑
"""

import os
import asyncio
import edge_tts


# ══════════════════════════════════════════════════════════════
# 單頁音檔生成（含重試與驗證）
# ══════════════════════════════════════════════════════════════

async def _generate_one(
    text: str,
    path: str,
    voice: str,
    retries: int = 3,
):
    """
    生成單頁音檔。
    - 自動重試最多 retries 次
    - 驗證檔案大小 > 1KB，過小視為損毀並重試
    - text 支援純文字或 SSML 格式
    """
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

    print(f"\n  [錯誤] {os.path.basename(path)} 已達最大重試次數，略過此頁")


# ══════════════════════════════════════════════════════════════
# 批次生成
# ══════════════════════════════════════════════════════════════

async def generate_audio_for_targets(
    scripts: list,
    target_indices: set,
    audio_dir: str,
    settings: dict,
) -> dict:
    """
    只為目標頁生成音檔，支援斷點續跑（已存在的音檔跳過）。

    回傳 {0-based-idx: audio_path} 的 dict。
    """
    voice      = settings.get("tts_voice", "zh-TW-HsiaoChenNeural")
    batch_size = settings.get("tts_batch_size", 20)

    targets   = sorted(target_indices) if target_indices else list(range(len(scripts)))
    audio_map = {}

    for batch_start in range(0, len(targets), batch_size):
        batch = targets[batch_start: batch_start + batch_size]
        tasks = []

        for idx in batch:
            audio_path = os.path.join(audio_dir, f"slide_{idx + 1:04d}.mp3")
            audio_map[idx] = audio_path

            if not os.path.exists(audio_path):
                tasks.append(_generate_one(scripts[idx], audio_path, voice))

        if tasks:
            await asyncio.gather(*tasks)

        done = min(batch_start + batch_size, len(targets))
        print(f"  TTS 進度：{done}/{len(targets)}", end="\r")

    print()
    return audio_map
