"""
phase2_compose.py
─────────────────
Phase 2：掃描 clips → 完整性檢查 → ffmpeg 串接成最終 MP4
        → 一鍵生成資通院驗收交付檔案（字幕檔、完整講稿、分段時間大綱）

使用方式：
  python phase2_compose.py
  python phase2_compose.py --output final.mp4
  python phase2_compose.py --pptx 教學簡報測試.pptx --output 教學簡報.mp4
  python phase2_compose.py --output final.mp4 --clean
"""

import os
import re
import sys
import glob
import shutil
import argparse
import yaml

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from core.video import compose_video
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


def scan_clips(clips_dir: str) -> list:
    """
    掃描 clips/ 下所有子資料夾，回傳按正確順序排列的 clip 路徑。
    順序：intro/ 優先，其餘按數字升序，資料夾內按頁碼升序。
    """
    if not os.path.isdir(clips_dir):
        print(f"[錯誤] clips 資料夾不存在：{clips_dir}")
        sys.exit(1)

    intro_folder     = None
    numbered_folders = []

    for name in os.listdir(clips_dir):
        full = os.path.join(clips_dir, name)
        if not os.path.isdir(full):
            continue
        if name.lower() == "intro":
            intro_folder = full
        elif re.match(r"^\d+$", name):
            numbered_folders.append((int(name), full))

    numbered_folders.sort(key=lambda x: x[0])

    ordered = []
    if intro_folder:
        ordered.append(("intro", intro_folder))
    for num, path in numbered_folders:
        ordered.append((str(num), path))

    all_clips = []
    for _, folder_path in ordered:
        clips = [
            f for f in os.listdir(folder_path)
            if f.lower().endswith(".mp4")
        ]
        clips.sort(key=lambda x: int("".join(filter(str.isdigit, x)) or 0))
        for clip in clips:
            all_clips.append(os.path.join(folder_path, clip))

    return all_clips


def check_clips(clips: list) -> bool:
    """檢查每個 clip 是否存在且非空，列出問題後回傳是否全部正常"""
    issues = []
    for path in clips:
        if not os.path.exists(path):
            issues.append(f"  [缺漏] {path}")
        elif os.path.getsize(path) == 0:
            issues.append(f"  [空檔] {path}")

    if issues:
        print(f"\n[警告] 發現 {len(issues)} 個問題：")
        for issue in issues:
            print(issue)
        return False
    return True


def print_clips_summary(clips: list):
    folder_counts = {}
    for path in clips:
        folder = os.path.basename(os.path.dirname(path))
        folder_counts[folder] = folder_counts.get(folder, 0) + 1

    print(f"\n  clips 摘要（共 {len(clips)} 個）：")
    for folder, count in folder_counts.items():
        print(f"    {folder}/  ->  {count} 個 clip")


def main():
    parser = argparse.ArgumentParser(description="Phase 2：clips -> 最終 MP4 與資通院交付物")
    parser.add_argument("--pptx",            default="",                 help="PPT 檔案路徑（用於決定隔離資料夾與生成字幕/大綱）")
    parser.add_argument("--clips",           default="clips",            help="clips 資料夾路徑（預設自動匹配 clips/<PPT檔名>）")
    parser.add_argument("--output",          default="",                 help="輸出影片檔名（預設依 PPT 檔名命名為 <PPT檔名>.mp4）")
    parser.add_argument("--output-dir",      default="",                 help="交付物輸出資料夾（預設依 PPT 檔名分流至 deliverables/<PPT檔名>）")
    parser.add_argument("--settings",        default="settings.yaml",    help="設定檔路徑")
    parser.add_argument("--clean",           action="store_true",        help="合成完後刪除 clips 資料夾")
    parser.add_argument("--no-deliverables", action="store_true",        help="不生成字幕與分段大綱交付物")
    args = parser.parse_args()

    pptx_path = os.path.abspath(args.pptx) if args.pptx else find_default_pptx()
    pptx_stem = get_pptx_stem(pptx_path) if pptx_path else ""

    # 自動解析 clips 資料夾：若使用者指定為預設 "clips" 且有 pptx_path
    if args.clips == "clips" and pptx_path:
        stem_clips = get_default_clips_dir(pptx_path, "clips")
        if os.path.isdir(stem_clips):
            clips_dir = stem_clips
        else:
            clips_dir = os.path.abspath(args.clips)
    else:
        clips_dir = os.path.abspath(args.clips)

    # 自動解析 output_dir
    if args.output_dir:
        output_dir = os.path.abspath(args.output_dir)
    elif pptx_path:
        output_dir = get_default_deliverables_dir(pptx_path, "deliverables")
    else:
        output_dir = os.path.abspath("deliverables")

    os.makedirs(output_dir, exist_ok=True)

    # 自動決定輸出影片檔名
    out_name = args.output.strip()
    if not out_name:
        out_name = f"{pptx_stem}.mp4" if pptx_stem else "final_output.mp4"

    if not os.path.dirname(out_name):
        output_video = os.path.join(output_dir, out_name)
    else:
        output_video = os.path.abspath(out_name)
        output_dir   = os.path.dirname(output_video)

    settings     = load_settings(args.settings, pptx_path)
    ffmpeg_path  = settings.get("ffmpeg_path", "ffmpeg")


    print("=" * 70)
    print("Phase 2：掃描 clips 資料夾...")
    print("=" * 70)
    clips = scan_clips(clips_dir)

    if not clips:
        print("[錯誤] 找不到任何 clip 檔案，請先執行 phase1_generate.py")
        sys.exit(1)

    print_clips_summary(clips)

    print("\n執行完整性檢查...")
    if not check_clips(clips):
        print("\n請先補齊缺漏的 clip 再執行合成。")
        print("提示：執行 phase1_generate.py --pages <缺漏頁碼> 補生成。")
        sys.exit(1)

    print("  [OK] 所有 clip 檢查通過")

    # 1. 影片串接
    compose_video(clips, output_video, clips_dir, ffmpeg_path)
    print(f"\n[完成] 影片串接成功！輸出原檔：{output_video}")

    # 2. 自動產出字幕、講稿與分段時間大綱
    if not args.no_deliverables:
        pptx_path = args.pptx or find_default_pptx()
        if pptx_path and os.path.exists(pptx_path):
            base_name = os.path.splitext(os.path.basename(output_video))[0]
            print(f"\n開始生成字幕檔、完整講稿與分段大綱...")
            export_all_deliverables(
                clips=clips,
                pptx_path=pptx_path,
                output_dir=output_dir,
                settings=settings,
                base_name=base_name,
            )
        else:
            print(f"\n[提示] 找不到對應的 PPTX 檔案，略過字幕與大綱生成。可使用 --pptx 指定路徑。")

    if args.clean:
        shutil.rmtree(clips_dir)
        print(f"   clips 資料夾已刪除：{clips_dir}")


if __name__ == "__main__":
    main()
