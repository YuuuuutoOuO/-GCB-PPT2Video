"""
generate_deliverables.py
────────────────────────
獨立交付物生成工具：
不需重新跑 TTS 配音或重新合成影片，直接掃描現有 clips 與 PPT，
秒級產出資通院驗收所需的：
  1. 完整講稿內容（.txt / .md）
  2. 影片字幕檔（.srt / .vtt）
  3. 分段大綱含對應影片時間點（幾分幾秒，.txt / .md / .csv）

使用方式：
  python generate_deliverables.py
  python generate_deliverables.py --pptx 教學簡報測試.pptx
  python generate_deliverables.py --pptx 教學簡報測試.pptx --output-dir 交付檔案
"""

import os
import sys
import glob
import argparse
import yaml

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from phase2_compose import scan_clips
from core.deliverables import export_all_deliverables
from core.utils import (
    load_settings, get_default_clips_dir,
    get_default_deliverables_dir, get_pptx_stem,
)


def find_default_pptx() -> str:
    candidates = ["教學簡報測試.pptx", "測試.pptx"]
    for c in candidates:
        if os.path.exists(c):
            return c
    all_pptx = glob.glob("*.pptx")
    return all_pptx[0] if all_pptx else ""


def main():
    parser = argparse.ArgumentParser(description="一鍵生成資通院驗收交付檔案（字幕、講稿、分段大綱時間點）")
    parser.add_argument("--pptx",       default="",              help="PPT 檔案路徑（預設自動尋找）")
    parser.add_argument("--clips",      default="clips",         help="clips 資料夾路徑（預設自動匹配 clips/<PPT檔名>）")
    parser.add_argument("--output-dir", default="",              help="交付物輸出資料夾（預設依 PPT 檔名分流至 deliverables/<PPT檔名>）")
    parser.add_argument("--settings",   default="settings.yaml", help="設定檔路徑")
    parser.add_argument("--name",       default="",              help="交付物基礎檔名（預設依 PPT 檔名命名）")
    args = parser.parse_args()

    pptx_path = os.path.abspath(args.pptx) if args.pptx else find_default_pptx()
    if not pptx_path or not os.path.exists(pptx_path):
        print(f"[錯誤] 找不到 PPT 檔案：{pptx_path}，請使用 --pptx 指定正確路徑。")
        sys.exit(1)

    pptx_stem = get_pptx_stem(pptx_path)

    # 自動解析 clips 資料夾：若使用者指定為預設 "clips"
    if args.clips == "clips":
        stem_clips = get_default_clips_dir(pptx_path, "clips")
        if os.path.isdir(stem_clips):
            clips_dir = stem_clips
        else:
            clips_dir = os.path.abspath(args.clips)
    else:
        clips_dir = os.path.abspath(args.clips)

    if not os.path.isdir(clips_dir):
        print(f"[錯誤] clips 資料夾不存在：{clips_dir}")
        sys.exit(1)

    # 自動解析 output_dir
    if args.output_dir:
        output_dir = os.path.abspath(args.output_dir)
    else:
        output_dir = get_default_deliverables_dir(pptx_path, "deliverables")

    os.makedirs(output_dir, exist_ok=True)

    settings = load_settings(args.settings, pptx_path)
    base_name = args.name.strip() or pptx_stem

    print("=" * 70)
    print("資通院交付物一鍵生成工具")
    print("=" * 70)
    print(f"  PPT 檔案   ：{pptx_path}")
    print(f"  clips 目錄 ：{clips_dir}")
    print(f"  輸出資料夾 ：{output_dir}")
    print(f"  交付物名稱 ：{base_name}")

    clips = scan_clips(clips_dir)
    if not clips:
        print("[錯誤] 找不到任何 clip 影片檔，請先執行 phase1_generate.py 生成片段。")
        sys.exit(1)

    print(f"  掃描到 clip 數量：{len(clips)} 個")

    export_all_deliverables(
        clips=clips,
        pptx_path=pptx_path,
        output_dir=output_dir,
        settings=settings,
        base_name=base_name,
    )



if __name__ == "__main__":
    main()
