"""
phase2_compose.py
─────────────────
Phase 2：掃描 clips → 完整性檢查 → ffmpeg 串接成最終 MP4

使用方式：
  python phase2_compose.py
  python phase2_compose.py --clips clips --output final.mp4
  python phase2_compose.py --output final.mp4 --clean
"""

import os
import re
import sys
import shutil
import argparse
import yaml

from core.video import compose_video


def load_settings(path: str) -> dict:
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


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
        print(f"\n⚠️  發現 {len(issues)} 個問題：")
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
        print(f"    {folder}/  →  {count} 個 clip")


def main():
    parser = argparse.ArgumentParser(description="Phase 2：clips → 最終 MP4")
    parser.add_argument("--clips",    default="clips",            help="clips 資料夾路徑")
    parser.add_argument("--output",   default="final_output.mp4", help="輸出影片檔名")
    parser.add_argument("--settings", default="settings.yaml",    help="設定檔路徑")
    parser.add_argument("--clean",    action="store_true",        help="合成完後刪除 clips 資料夾")
    args = parser.parse_args()

    clips_dir    = os.path.abspath(args.clips)
    output_video = os.path.abspath(args.output)
    settings     = load_settings(args.settings)
    ffmpeg_path  = settings.get("ffmpeg_path", "ffmpeg")

    print("Phase 2：掃描 clips 資料夾...")
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

    print("  ✅ 所有 clip 檢查通過")

    compose_video(clips, output_video, clips_dir, ffmpeg_path)

    print(f"\n✅ Phase 2 完成！輸出：{output_video}")

    if args.clean:
        shutil.rmtree(clips_dir)
        print(f"   clips 資料夾已刪除：{clips_dir}")


if __name__ == "__main__":
    main()
