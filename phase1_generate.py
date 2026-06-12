"""
phase1_generate.py
──────────────────
Phase 1：解析 PPT → 組合講稿 → TTS 配音 → 合成單頁 MP4

使用方式：
  python phase1_generate.py --pptx file.pptx
  python phase1_generate.py --pptx file.pptx --pages 1-3,5
  python phase1_generate.py --pptx file.pptx --range 58,60-65
  python phase1_generate.py --pptx file.pptx --pages 1-3 --range 58-60
  python phase1_generate.py --pptx file.pptx --workers 8   （調整平行數）
  python phase1_generate.py --pptx file.pptx --no-gpu      （強制 CPU 編碼）
"""

import os
import asyncio
import tempfile
import argparse
import yaml

from core.parser  import parse_all_slides
from core.script  import build_all_scripts
from core.tts     import generate_audio_for_targets
from core.video   import (
    export_slides_to_images, get_sorted_images,
    get_clip_path, create_clips_parallel,
)
from core.utils   import resolve_target_indices, confirm_dual_mode


def load_settings(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def write_scripts_txt(scripts, parsed_list, target_indices, output_path):
    """只輸出本次處理頁的講稿（scripts 為 list[list[str]]）"""
    targets = sorted(target_indices) if target_indices else list(range(len(scripts)))
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("本次處理頁講稿預覽\n")
        f.write("=" * 60 + "\n\n")
        for idx in targets:
            parsed  = parsed_list[idx]
            preview = " / ".join(scripts[idx])[:200] if scripts[idx] else "（空白頁）"
            f.write(f"【第 {idx + 1} 頁】\n")
            f.write(f"  標題   ：{parsed['title']}\n")
            f.write(f"  副標題 ：{parsed['subtitle']}\n")
            f.write(f"  講稿   ：{preview}\n\n")
    print(f"  講稿已輸出至：{output_path}")


def main():
    parser = argparse.ArgumentParser(description="Phase 1：PPT → 單頁 MP4 clips")
    parser.add_argument("--pptx",     required=True,           help="PPT 檔案路徑")
    parser.add_argument("--settings", default="settings.yaml", help="設定檔路徑")
    parser.add_argument("--clips",    default="clips",         help="clips 輸出資料夾")
    parser.add_argument("--pages",    default="",              help="指定頁碼，如 1-3,5,10")
    parser.add_argument("--range",    default="",              help="指定標題編號，如 58,60-65")
    parser.add_argument("--workers",  type=int, default=4,     help="Phase C 平行合成數（預設 4）")
    parser.add_argument("--no-gpu",   action="store_true",     help="停用 GPU 編碼，改用 CPU libx264")
    args = parser.parse_args()

    pptx_path = os.path.abspath(args.pptx)
    clips_dir = os.path.abspath(args.clips)
    os.makedirs(clips_dir, exist_ok=True)

    print("載入設定檔...")
    settings    = load_settings(args.settings)
    fps         = settings.get("video_fps", 12)
    dpi         = settings.get("export_dpi", 200)
    ffmpeg_path = settings.get("ffmpeg_path", "ffmpeg")
    use_gpu     = not args.no_gpu

    print(f"  ffmpeg：{ffmpeg_path}")
    print(f"  編碼  ：{'NVIDIA GPU (h264_nvenc)' if use_gpu else 'CPU (libx264)'}")
    print(f"  平行數：{args.workers}")

    print("\n解析 PPT 結構...")
    parsed_list = parse_all_slides(pptx_path)
    print(f"  共 {len(parsed_list)} 頁")

    print("組合講稿...")
    scripts = build_all_scripts(parsed_list, settings)

    target_indices = resolve_target_indices(parsed_list, args.pages, args.range)
    if args.pages and args.range:
        confirm_dual_mode(parsed_list, args.pages, args.range, target_indices)

    targets = sorted(target_indices) if target_indices else list(range(len(scripts)))
    print(f"\n本次處理：{len(targets)} 頁")

    scripts_txt_path = os.path.join(clips_dir, "scripts.txt")
    write_scripts_txt(scripts, parsed_list, target_indices, scripts_txt_path)

    with tempfile.TemporaryDirectory() as tmp_dir:
        image_dir = os.path.join(tmp_dir, "images")
        audio_dir = os.path.join(tmp_dir, "audio")
        os.makedirs(image_dir, exist_ok=True)
        os.makedirs(audio_dir, exist_ok=True)

        print("\nPhase A：匯出投影片截圖（暫存）...")
        export_slides_to_images(pptx_path, image_dir, dpi=dpi)
        images = get_sorted_images(image_dir)

        print("\nPhase B：生成 AI 配音（暫存）...")
        audio_map = asyncio.run(
            generate_audio_for_targets(scripts, target_indices, audio_dir, settings)
        )

        # ── 過濾：跳過截圖缺失、音檔缺失、已存在的頁 ──────────────
        print("\nPhase C：合成單頁 MP4（平行處理）...")
        todo = []
        skipped = 0

        for idx in targets:
            if idx >= len(images):
                print(f"  [警告] 第 {idx + 1} 頁找不到截圖，跳過")
                continue

            audio_path = audio_map.get(idx, "")
            if not audio_path or not os.path.exists(audio_path):
                print(f"  [警告] 第 {idx + 1} 頁音檔不存在，跳過")
                continue

            clip_path = get_clip_path(parsed_list[idx], idx + 1, clips_dir)
            if os.path.exists(clip_path):
                skipped += 1
                continue

            todo.append(idx)

        if skipped:
            print(f"  已跳過 {skipped} 頁（clip 已存在）")

        audio_missing = len(targets) - len(todo) - skipped
        if audio_missing > 0:
            print(f"  ⚠️  {audio_missing} 頁音檔不存在，請確認 Phase B 是否成功")

        if todo:
            done_count, fail_count = create_clips_parallel(
                todo=todo,
                images=images,
                image_dir=image_dir,
                audio_map=audio_map,
                parsed_list=parsed_list,
                clips_dir=clips_dir,
                fps=fps,
                ffmpeg_path=ffmpeg_path,
                use_gpu=use_gpu,
                max_workers=args.workers,
            )
        else:
            done_count, fail_count = 0, 0
            if skipped == len(targets):
                print("  所有頁面皆已存在，無需重新合成。")
            else:
                print("  ⚠️  無可合成的頁面（音檔缺失或截圖缺失），請確認 Phase B 輸出。")

        print(f"\n✅ Phase 1 完成！")
        print(f"   成功合成：{done_count} 頁")
        if fail_count > 0:
            print(f"   ⚠️  失敗頁數：{fail_count} 頁（請補跑失敗頁）")
        print(f"   clips 位置：{clips_dir}")
        print(f"   講稿預覽  ：{scripts_txt_path}")


if __name__ == "__main__":
    main()