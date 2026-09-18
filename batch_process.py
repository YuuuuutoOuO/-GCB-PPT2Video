"""
batch_process.py
────────────────
PPT 批次全自動轉換工具（支援 20+ 份 PPT 一鍵排隊轉換）：
  1. 自動掃描指定資料夾內所有 PPTX 檔案（排除 ~$ 暫存檔）
  2. 依據 settings.yaml 自動套用每份 PPT 的專屬設定（如 notes_slides）
  3. 各 PPT 自動進行目錄物理隔離（clips/<PPT檔名>/ 與 deliverables/<PPT檔名>/）
  4. 支援 --skip-existing：已完成的 PPT 自動跳過
  5. 支援 --clean：每份 PPT 生成完畢後自動清理 clips 暫存檔以節省磁碟空間
  6. 輸出詳細執行進度與最終彙總報表

使用方式：
  python batch_process.py
  python batch_process.py --dir ./my_pptx_folder
  python batch_process.py --dir ./my_pptx_folder --skip-existing --clean
  python batch_process.py --dir ./my_pptx_folder --workers 8
"""

import os
import sys
import glob
import time
import shutil
import argparse
import subprocess
from datetime import datetime

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from core.utils import (
    load_settings, get_pptx_stem,
    get_default_clips_dir, get_default_deliverables_dir,
)


def format_duration(seconds: float) -> str:
    """將秒數格式化為 幾時幾分幾秒 或 幾分幾秒"""
    sec = max(0, int(round(seconds)))
    hrs = sec // 3600
    mins = (sec % 3600) // 60
    secs = sec % 60
    if hrs > 0:
        return f"{hrs} 小時 {mins:02d} 分 {secs:02d} 秒"
    if mins > 0:
        return f"{mins} 分 {secs:02d} 秒"
    return f"{secs} 秒"


def find_pptx_files(directory: str) -> list[str]:
    """掃描指定目錄下所有 .pptx 檔案，排除 office 暫存檔 (~$ 開頭)"""
    abs_dir = os.path.abspath(directory)
    if not os.path.isdir(abs_dir):
        return []

    pattern = os.path.join(abs_dir, "*.pptx")
    files = glob.glob(pattern)
    valid_files = [
        f for f in files
        if not os.path.basename(f).startswith("~$")
    ]
    valid_files.sort(key=lambda x: os.path.basename(x).lower())
    return valid_files


def run_single_ppt(
    pptx_path: str,
    settings_path: str,
    base_clips_dir: str,
    base_output_dir: str,
    workers: int,
    use_gpu: bool,
    clean_clips: bool,
    notes_slides: str = "",
) -> bool:
    """執行單一 PPT 的完整 Phase 1 + Phase 2 流程"""
    stem = get_pptx_stem(pptx_path)
    clips_dir = get_default_clips_dir(pptx_path, base_clips_dir)
    output_dir = get_default_deliverables_dir(pptx_path, base_output_dir)
    output_video = os.path.join(output_dir, f"{stem}.mp4")

    os.makedirs(clips_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    # 1. 組合 Phase 1 指令
    cmd_p1 = [
        sys.executable, "-u", "phase1_generate.py",
        "--pptx", pptx_path,
        "--settings", settings_path,
        "--clips", clips_dir,
        "--workers", str(workers),
    ]
    if notes_slides.strip():
        cmd_p1.extend(["--notes-slides", notes_slides.strip()])
    if not use_gpu:
        cmd_p1.append("--no-gpu")

    print(f"\n[Phase 1] 開始生成 clips 片段與講稿...")
    proc_p1 = subprocess.run(cmd_p1)
    if proc_p1.returncode != 0:
        print(f"❌ [錯誤] Phase 1 執行失敗（結束代碼：{proc_p1.returncode}）")
        return False

    # 2. 組合 Phase 2 指令
    cmd_p2 = [
        sys.executable, "-u", "phase2_compose.py",
        "--pptx", pptx_path,
        "--settings", settings_path,
        "--clips", clips_dir,
        "--output", output_video,
        "--output-dir", output_dir,
    ]

    print(f"\n[Phase 2] 開始串接影片與生成資通院交付檔案...")
    proc_p2 = subprocess.run(cmd_p2)
    if proc_p2.returncode != 0:
        print(f"❌ [錯誤] Phase 2 執行失敗（結束代碼：{proc_p2.returncode}）")
        return False

    # 3. 若勾選 clean，清理 clips 目錄節省硬碟空間
    if clean_clips and os.path.exists(clips_dir):
        try:
            shutil.rmtree(clips_dir)
            print(f"🧹 [清理] clips 暫存資料夾已刪除：{clips_dir}")
        except Exception as e:
            print(f"⚠️ [警告] 刪除 clips 暫存資料夾失敗：{e}")

    return True


def main():
    parser = argparse.ArgumentParser(description="PPT 批次全自動轉換工具（支援 20+ 份 PPT 依序排程轉換）")
    parser.add_argument("--dir",           default=".",              help="PPT 檔案所在資料夾（預設目前目錄）")
    parser.add_argument("--settings",      default="settings.yaml",  help="設定檔路徑（預設 settings.yaml）")
    parser.add_argument("--clips-dir",     default="clips",          help="clips 輸出基礎資料夾（預設 clips）")
    parser.add_argument("--output-dir",    default="deliverables",   help="交付物輸出基礎資料夾（預設 deliverables）")
    parser.add_argument("--skip-existing", action="store_true",      help="若該 PPT 之交付影片已存在則自動跳過")
    parser.add_argument("--clean",         action="store_true",      help="每份 PPT 完成後自動刪除 clips 暫存檔以節省磁碟空間")
    parser.add_argument("--workers",       type=int, default=4,      help="Phase C 平行合成數（預設 4）")
    parser.add_argument("--no-gpu",        action="store_true",      help="停用 GPU 編碼，改用 CPU libx264")
    args = parser.parse_args()

    input_dir = os.path.abspath(args.dir)
    print("=" * 75)
    print("🚀 PPT 批次全自動轉換排程系統")
    print("=" * 75)
    print(f"  掃描資料夾     ：{input_dir}")
    print(f"  設定檔路徑     ：{os.path.abspath(args.settings)}")
    print(f"  交付物總目錄   ：{os.path.abspath(args.output_dir)}")
    print(f"  跳過已完成 PPT ：{'是 (啟用)' if args.skip_existing else '否'}")
    print(f"  完成後清除暫存 ：{'是 (啟用 --clean 節省硬碟空間)' if args.clean else '否'}")
    print(f"  編碼加速       ：{'CPU (libx264)' if args.no_gpu else 'NVIDIA GPU (h264_nvenc)'}")
    print(f"  平行合成數     ：{args.workers}")

    pptx_files = find_pptx_files(input_dir)
    total_count = len(pptx_files)

    if total_count == 0:
        print(f"\n[錯誤] 在資料夾【{input_dir}】中找不到任何 .pptx 簡報檔案！")
        sys.exit(1)

    print(f"\n共掃描到 {total_count} 個 PPT 簡報檔案：")
    for idx, f in enumerate(pptx_files, start=1):
        stem = get_pptx_stem(f)
        cfg = load_settings(args.settings, f)
        notes = cfg.get("notes_slides", [])
        print(f"  [{idx:02d}/{total_count:02d}] {os.path.basename(f)} (專屬 notes_slides: {notes})")

    print("\n" + "=" * 75)
    print("開始依序執行批次處理...")
    print("=" * 75)

    start_all_time = time.time()
    results_success = []
    results_failed = []
    results_skipped = []

    for idx, pptx_path in enumerate(pptx_files, start=1):
        stem = get_pptx_stem(pptx_path)
        ppt_name = os.path.basename(pptx_path)
        ppt_output_dir = get_default_deliverables_dir(pptx_path, args.output_dir)
        target_video = os.path.join(ppt_output_dir, f"{stem}.mp4")

        print(f"\n{'#'*75}")
        print(f"▶ 進度 [{idx}/{total_count}] 正在處理：{ppt_name}")
        print(f"  檔案路徑：{pptx_path}")
        print(f"  輸出目標：{target_video}")
        print(f"{'#'*75}")

        # 檢查是否跳過已存在的 PPT
        if args.skip_existing and os.path.exists(target_video) and os.path.getsize(target_video) > 1024:
            print(f"⏩ [略過] 偵測到目標影片已存在且非空檔，自動跳過：{target_video}")
            results_skipped.append((ppt_name, "影片已存在"))
            continue

        start_single_time = time.time()
        success = run_single_ppt(
            pptx_path=pptx_path,
            settings_path=args.settings,
            base_clips_dir=args.clips_dir,
            base_output_dir=args.output_dir,
            workers=args.workers,
            use_gpu=not args.no_gpu,
            clean_clips=args.clean,
        )
        elapsed_single = time.time() - start_single_time

        if success:
            print(f"\n✅ [{idx}/{total_count}] 【{ppt_name}】處理完成！耗時：{format_duration(elapsed_single)}")
            results_success.append((ppt_name, elapsed_single))
        else:
            print(f"\n❌ [{idx}/{total_count}] 【{ppt_name}】處理失敗！耗時：{format_duration(elapsed_single)}")
            results_failed.append((ppt_name, "處理過程出錯"))

    total_elapsed = time.time() - start_all_time

    # 彙總報表
    print("\n" + "=" * 75)
    print("📊 批次處理彙總報告 (Batch Summary Report)")
    print("=" * 75)
    print(f"  總處理簡報數 ：{total_count} 份")
    print(f"  ✅ 成功數量  ：{len(results_success)} 份")
    print(f"  ⏩ 跳過數量  ：{len(results_skipped)} 份")
    print(f"  ❌ 失敗數量  ：{len(results_failed)} 份")
    print(f"  ⏱️ 總耗時     ：{format_duration(total_elapsed)}")

    if results_success:
        print("\n  [成功清單]：")
        for name, dur in results_success:
            print(f"    • {name}（耗時 {format_duration(dur)}）")

    if results_skipped:
        print("\n  [跳過清單]：")
        for name, reason in results_skipped:
            print(f"    • {name}（{reason}）")

    if results_failed:
        print("\n  [失敗清單]：")
        for name, reason in results_failed:
            print(f"    • {name}（{reason}）")

    print("=" * 75 + "\n")


if __name__ == "__main__":
    main()
