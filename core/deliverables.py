"""
core/deliverables.py
───────────────────
資通院交付物匯出模組（完全符合資通院「繳交說明.docx」規範）：
  1. 影片時長與時間軸元資料統計 (ffprobe 平行提取)
  2. 字幕檔產出 (.srt, .vtt，依自然句/逗號斷句，對齊「字幕範例.srt」)
  3. MP4 章節檔 (.txt，FFMETADATA1 格式，對齊「大綱_MP4章節範例.txt」)
  4. 完整講稿內容匯出 (.txt, .md)
  5. 分段大綱時間點清單與 Excel 審查表 (.csv, .md, .txt)
"""

import os
import re
import csv
import shutil
import subprocess
from concurrent.futures import ThreadPoolExecutor
from typing import Optional

from core.utils import extract_title_number


# ══════════════════════════════════════════════════════════════
# 時間格式化工具
# ══════════════════════════════════════════════════════════════

def format_seconds_to_srt(sec: float) -> str:
    """轉換秒數為 SRT 時間戳格式：00:01:23,450"""
    sec = max(0.0, sec)
    hrs = int(sec // 3600)
    mins = int((sec % 3600) // 60)
    secs = int(sec % 60)
    millis = int(round((sec - int(sec)) * 1000))
    if millis >= 1000:
        secs += 1
        millis = 0
    return f"{hrs:02d}:{mins:02d}:{secs:02d},{millis:03d}"


def format_seconds_to_vtt(sec: float) -> str:
    """轉換秒數為 WebVTT 時間戳格式：00:01:23.450"""
    sec = max(0.0, sec)
    hrs = int(sec // 3600)
    mins = int((sec % 3600) // 60)
    secs = int(sec % 60)
    millis = int(round((sec - int(sec)) * 1000))
    if millis >= 1000:
        secs += 1
        millis = 0
    return f"{hrs:02d}:{mins:02d}:{secs:02d}.{millis:03d}"


def format_seconds_to_hms(sec: float) -> str:
    """轉換秒數為標準 HH:MM:SS"""
    sec = max(0.0, round(sec))
    hrs = int(sec // 3600)
    mins = int((sec % 3600) // 60)
    secs = int(sec % 60)
    return f"{hrs:02d}:{mins:02d}:{secs:02d}"


def format_seconds_to_chinese_time(sec: float) -> str:
    """轉換秒數為中文易讀格式（幾分幾秒 / 幾時幾分幾秒）"""
    sec = max(0.0, round(sec))
    hrs = int(sec // 3600)
    mins = int((sec % 3600) // 60)
    secs = int(sec % 60)
    if hrs > 0:
        return f"{hrs}時{mins:02d}分{secs:02d}秒"
    return f"{mins:02d}分{secs:02d}秒"


# ══════════════════════════════════════════════════════════════
# 時長讀取與時間軸統計
# ══════════════════════════════════════════════════════════════

def _find_ffprobe(ffmpeg_path: str = "ffmpeg") -> str:
    """依據 ffmpeg 路徑推導或尋找 ffprobe 執行檔"""
    if os.path.isabs(ffmpeg_path) or "/" in ffmpeg_path or "\\" in ffmpeg_path:
        base_dir = os.path.dirname(ffmpeg_path)
        candidate = os.path.join(base_dir, "ffprobe.exe" if os.name == "nt" else "ffprobe")
        if os.path.exists(candidate):
            return candidate
    local_ffprobe = os.path.join(".", "ffmpeg", "ffprobe.exe")
    if os.path.exists(local_ffprobe):
        return os.path.abspath(local_ffprobe)
    return "ffprobe"


def get_single_clip_duration(clip_path: str, ffprobe_path: str = "ffprobe") -> float:
    """使用 ffprobe 獲取單個 clip 的精確時長（秒）"""
    if not os.path.exists(clip_path) or os.path.getsize(clip_path) == 0:
        return 0.0
    try:
        out = subprocess.check_output([
            ffprobe_path, "-v", "error",
            "-show_entries", "format=duration",
            "-of", "default=noprint_wrappers=1:nokey=1",
            clip_path,
        ], stderr=subprocess.DEVNULL)
        return float(out.decode().strip())
    except Exception:
        return 0.0


def extract_all_clip_durations(
    clips: list[str],
    ffmpeg_path: str = "ffmpeg",
    max_workers: int = 16,
) -> dict[str, float]:
    """多執行緒平行提取所有 clip 影片時長"""
    ffprobe_path = _find_ffprobe(ffmpeg_path)
    durations = {}

    def _task(clip):
        return clip, get_single_clip_duration(clip, ffprobe_path)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        results = executor.map(_task, clips)
        for clip, dur in results:
            durations[clip] = dur

    return durations


def build_timeline_metadata(
    clips: list[str],
    parsed_list: list[dict],
    scripts: list[list[str]],
    durations: dict[str, float],
) -> list[dict]:
    """
    整合 clips、PPT 解析結構、講稿與時長，生成完整時間軸元資料列表。
    """
    timeline = []
    current_time = 0.0

    for clip in clips:
        # 從檔名抓取 1-based 頁碼
        filename = os.path.basename(clip)
        match = re.search(r"page_(\d+)", filename)
        page_1based = int(match.group(1)) if match else 0
        page_idx = page_1based - 1

        duration = durations.get(clip, 0.0)
        start_time = current_time
        end_time = current_time + duration
        current_time = end_time

        parsed = parsed_list[page_idx] if 0 <= page_idx < len(parsed_list) else {}
        slide_scripts = scripts[page_idx] if 0 <= page_idx < len(scripts) else []

        title = parsed.get("title", "")
        subtitle = parsed.get("subtitle", "")
        title_num = extract_title_number(title)

        timeline.append({
            "clip_path": clip,
            "page_1based": page_1based,
            "page_idx": page_idx,
            "title": title,
            "title_num": title_num,
            "subtitle": subtitle,
            "scripts": slide_scripts,
            "start_sec": start_time,
            "end_sec": end_time,
            "duration_sec": duration,
        })

    return timeline


# ══════════════════════════════════════════════════════════════
# 1. 字幕檔生成 (.srt / .vtt) — 符合「字幕範例.srt」自然斷句
# ══════════════════════════════════════════════════════════════

def _split_slide_scripts_into_cues(paragraphs: list[str], total_duration: float, start_offset: float):
    """
    將單頁段落文字依照標點符號與子句分割為短句字幕，並依字數比例精確計算起訖時間戳。
    每句長度約 12~28 字，閱讀體驗舒適。
    """
    clauses = []
    for p in paragraphs:
        p = p.strip()
        if not p:
            continue
        # 依句號、問號、驚嘆號、分號、換行分段
        parts = re.split(r'([。！？；\n]+)', p)
        cur = ""
        for part in parts:
            if not part:
                continue
            if re.match(r'^[。！？；\n]+$', part):
                cur += part
                if cur.strip():
                    clauses.append(cur.strip())
                cur = ""
            else:
                # 若單句過長（超過 25 字）且含有逗號/頓號，依逗號進一步細切
                if len(cur) + len(part) > 25 and ('，' in part or '、' in part or ',' in part):
                    sub_parts = re.split(r'([，,、])', part)
                    for sp in sub_parts:
                        if not sp:
                            continue
                        if sp in [',', '，', '、']:
                            cur += sp
                            if len(cur) >= 8:
                                clauses.append(cur.strip())
                                cur = ""
                        else:
                            cur += sp
                else:
                    cur += part
        if cur.strip():
            clauses.append(cur.strip())

    if not clauses:
        return []

    total_chars = sum(max(1, len(c)) for c in clauses)
    cues = []
    curr_t = start_offset
    for idx, c in enumerate(clauses):
        ratio = max(1, len(c)) / total_chars
        dur = total_duration * ratio
        end_t = curr_t + dur
        if idx == len(clauses) - 1:
            end_t = start_offset + total_duration
        cues.append((curr_t, end_t, c))
        curr_t = end_t

    return cues


def export_subtitles(
    timeline: list[dict],
    srt_path: str,
    vtt_path: Optional[str] = None,
):
    """
    生成標準 SRT 與 WebVTT 字幕檔。
    """
    srt_lines = []
    vtt_lines = ["WEBVTT", ""]
    cue_index = 1

    for item in timeline:
        scripts = item["scripts"]
        start_sec = item["start_sec"]
        duration_sec = item["duration_sec"]

        if duration_sec <= 0.01 or not scripts:
            continue

        cues = _split_slide_scripts_into_cues(scripts, duration_sec, start_sec)
        for seg_start, seg_end, text in cues:
            text = text.strip()
            if not text:
                continue

            # SRT 格式
            srt_lines.append(str(cue_index))
            srt_lines.append(f"{format_seconds_to_srt(seg_start)} --> {format_seconds_to_srt(seg_end)}")
            srt_lines.append(text)
            srt_lines.append("")

            # VTT 格式
            vtt_lines.append(str(cue_index))
            vtt_lines.append(f"{format_seconds_to_vtt(seg_start)} --> {format_seconds_to_vtt(seg_end)}")
            vtt_lines.append(text)
            vtt_lines.append("")

            cue_index += 1

    os.makedirs(os.path.dirname(os.path.abspath(srt_path)), exist_ok=True)
    with open(srt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(srt_lines))

    if vtt_path:
        os.makedirs(os.path.dirname(os.path.abspath(vtt_path)), exist_ok=True)
        with open(vtt_path, "w", encoding="utf-8") as f:
            f.write("\n".join(vtt_lines))


# ══════════════════════════════════════════════════════════════
# 2. 章節檔生成 (.txt) — 100% 符合資通院「FFMETADATA1」格式
# ══════════════════════════════════════════════════════════════

def export_ffmetadata_chapters(
    timeline: list[dict],
    ffmetadata_path: str,
):
    """
    匯出資通院指定的 FFMETADATA1 章節大綱檔（.txt）。
    格式範例（完全同「大綱_MP4章節範例.txt」）：
      ;FFMETADATA1
      [CHAPTER]
      TIMEBASE=1/1000
      START=0
      END=15000
      title=開場引言
    """
    if not timeline:
        return

    chapters = []
    current_chapter = None

    for item in timeline:
        title = item["title"] or item["subtitle"] or "簡報說明"
        if not current_chapter or current_chapter["title"] != title:
            if current_chapter:
                current_chapter["end_ms"] = int(round(item["start_sec"] * 1000))
                chapters.append(current_chapter)
            current_chapter = {
                "title": title,
                "start_ms": int(round(item["start_sec"] * 1000)),
                "end_ms": int(round(item["end_sec"] * 1000)),
            }
        else:
            current_chapter["end_ms"] = int(round(item["end_sec"] * 1000))

    if current_chapter:
        chapters.append(current_chapter)

    lines = [";FFMETADATA1"]
    for ch in chapters:
        lines.append("[CHAPTER]")
        lines.append("TIMEBASE=1/1000")
        lines.append(f"START={ch['start_ms']}")
        lines.append(f"END={ch['end_ms']}")
        lines.append(f"title={ch['title']}")
        lines.append("")

    os.makedirs(os.path.dirname(os.path.abspath(ffmetadata_path)), exist_ok=True)
    with open(ffmetadata_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


# ══════════════════════════════════════════════════════════════
# 3. 完整講稿內容匯出 (.txt / .md)
# ══════════════════════════════════════════════════════════════

def export_full_transcript(
    timeline: list[dict],
    txt_path: str,
    md_path: Optional[str] = None,
):
    """
    匯出每頁投影片完整無截斷的講稿內容文件。
    """
    txt_lines = []
    txt_lines.append("【教學簡報與影片 — 完整講稿內容】")
    txt_lines.append("=" * 70)
    txt_lines.append(f"總頁數：{len(timeline)} 頁")
    if timeline:
        total_time = timeline[-1]["end_sec"]
        txt_lines.append(f"影片總長度：{format_seconds_to_hms(total_time)} ({format_seconds_to_chinese_time(total_time)})")
    txt_lines.append("=" * 70 + "\n")

    md_lines = []
    md_lines.append("# 教學簡報與影片 — 完整講稿內容\n")
    if timeline:
        total_time = timeline[-1]["end_sec"]
        md_lines.append(f"> **總頁數**：{len(timeline)} 頁 ｜ **影片總長度**：`{format_seconds_to_hms(total_time)}` ({format_seconds_to_chinese_time(total_time)})\n")
    md_lines.append("---\n")

    for item in timeline:
        page = item["page_1based"]
        title = item["title"] or "（無大標題）"
        subtitle = item["subtitle"] or "（無副標題）"
        time_display = f"{format_seconds_to_hms(item['start_sec'])} ~ {format_seconds_to_hms(item['end_sec'])} ({format_seconds_to_chinese_time(item['duration_sec'])})"

        txt_lines.append(f"【第 {page} 頁】 時間點：{time_display}")
        txt_lines.append(f"  大綱標題 ：{title}")
        txt_lines.append(f"  副標題   ：{subtitle}")
        txt_lines.append("  講稿內容 ：")
        if item["scripts"]:
            for s in item["scripts"]:
                txt_lines.append(f"    {s}")
        else:
            txt_lines.append("    （無配音/空白頁）")
        txt_lines.append("-" * 70 + "\n")

        md_lines.append(f"### 第 {page} 頁：{title} - {subtitle}\n")
        md_lines.append(f"- **影片時間點**：`{format_seconds_to_hms(item['start_sec'])}` ~ `{format_seconds_to_hms(item['end_sec'])}`（時長 `{format_seconds_to_chinese_time(item['duration_sec'])}`）")
        md_lines.append(f"- **大綱項目**：{title}")
        md_lines.append(f"- **分段副標**：{subtitle}\n")
        md_lines.append("**講稿內容**：\n")
        if item["scripts"]:
            for s in item["scripts"]:
                md_lines.append(f"> {s}\n")
        else:
            md_lines.append("> *（無配音/空白頁）*\n")
        md_lines.append("\n---\n")

    os.makedirs(os.path.dirname(os.path.abspath(txt_path)), exist_ok=True)
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(txt_lines))

    if md_path:
        os.makedirs(os.path.dirname(os.path.abspath(md_path)), exist_ok=True)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write("\n".join(md_lines))


# ══════════════════════════════════════════════════════════════
# 4. 分段大綱審查表 (.csv, .md, .txt)
# ══════════════════════════════════════════════════════════════

def export_chapter_outline(
    timeline: list[dict],
    txt_path: str,
    md_path: Optional[str] = None,
    csv_path: Optional[str] = None,
):
    """
    匯出分段大綱與影片對應時間點（供評委審查與核對）。
    """
    if not timeline:
        return

    total_time = timeline[-1]["end_sec"]

    # 整理章節分組
    chapters = []
    current_chapter = None

    for item in timeline:
        title = item["title"] or "前言/簡報說明"
        if not current_chapter or current_chapter["title"] != title:
            if current_chapter:
                current_chapter["end_sec"] = item["start_sec"]
                current_chapter["duration_sec"] = current_chapter["end_sec"] - current_chapter["start_sec"]
                chapters.append(current_chapter)
            current_chapter = {
                "title": title,
                "title_num": item["title_num"],
                "start_page": item["page_1based"],
                "end_page": item["page_1based"],
                "start_sec": item["start_sec"],
                "end_sec": item["end_sec"],
                "duration_sec": item["duration_sec"],
                "slides": [item],
            }
        else:
            current_chapter["end_page"] = item["page_1based"]
            current_chapter["end_sec"] = item["end_sec"]
            current_chapter["duration_sec"] = current_chapter["end_sec"] - current_chapter["start_sec"]
            current_chapter["slides"].append(item)

    if current_chapter:
        chapters.append(current_chapter)

    # TXT 易讀版
    txt_lines = []
    txt_lines.append("【教學簡報與影片 — 分段大綱與時間點清單】")
    txt_lines.append("=" * 70)
    txt_lines.append(f"影片總時長：{format_seconds_to_hms(total_time)} ({format_seconds_to_chinese_time(total_time)})")
    txt_lines.append(f"總章節數  ：共 {len(chapters)} 個主要單元 / {len(timeline)} 頁投影片")
    txt_lines.append("=" * 70 + "\n")

    txt_lines.append("一、主要章節大綱時間索引 (幾分幾秒)")
    txt_lines.append("-" * 70)
    for ch in chapters:
        time_str = format_seconds_to_hms(ch["start_sec"])
        chinese_time = format_seconds_to_chinese_time(ch["start_sec"])
        dur_str = format_seconds_to_chinese_time(ch["duration_sec"])
        page_str = f"第 {ch['start_page']}~{ch['end_page']} 頁" if ch['start_page'] != ch['end_page'] else f"第 {ch['start_page']} 頁"
        txt_lines.append(f"  [{time_str}] ({chinese_time}) {ch['title']} （{page_str}，時長 {dur_str}）")
        for s in ch["slides"]:
            sub = s["subtitle"] or "內容說明"
            s_time = format_seconds_to_hms(s["start_sec"])
            s_chinese = format_seconds_to_chinese_time(s["start_sec"])
            txt_lines.append(f"      • [{s_time}] ({s_chinese}) 第 {s['page_1based']:03d} 頁：{sub}")
        txt_lines.append("")

    txt_lines.append("\n" + "=" * 70)
    txt_lines.append("二、投影片詳細分段清單")
    txt_lines.append("-" * 70)
    for item in timeline:
        p = item["page_1based"]
        start_hms = format_seconds_to_hms(item["start_sec"])
        start_cn = format_seconds_to_chinese_time(item["start_sec"])
        dur_cn = format_seconds_to_chinese_time(item["duration_sec"])
        title = item["title"] or "（無大標題）"
        subtitle = item["subtitle"] or "（無副標題）"
        txt_lines.append(f"第 {p:03d} 頁 | 開始時間：{start_hms} ({start_cn}) | 時長：{dur_cn} | 大綱：{title} - {subtitle}")

    os.makedirs(os.path.dirname(os.path.abspath(txt_path)), exist_ok=True)
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(txt_lines))

    # Markdown 表格版
    if md_path:
        md_lines = []
        md_lines.append("# 教學簡報與影片 — 分段大綱與時間點\n")
        md_lines.append(f"> **影片總時長**：`{format_seconds_to_hms(total_time)}` ({format_seconds_to_chinese_time(total_time)}) ｜ **總單元數**：{len(chapters)} 個 ｜ **總頁數**：{len(timeline)} 頁\n")
        md_lines.append("## 一、主要單元大綱章節清單\n")
        md_lines.append("| 項目/章節 | 大綱名稱 | 頁碼範圍 | 開始時間點 (幾分幾秒) | 單元時長 |")
        md_lines.append("|:---:|:---|:---:|:---:|:---:|")
        for idx, ch in enumerate(chapters, start=1):
            page_str = f"P.{ch['start_page']}~P.{ch['end_page']}" if ch['start_page'] != ch['end_page'] else f"P.{ch['start_page']}"
            md_lines.append(
                f"| {idx} | **{ch['title']}** | {page_str} | `{format_seconds_to_hms(ch['start_sec'])}` ({format_seconds_to_chinese_time(ch['start_sec'])}) | {format_seconds_to_chinese_time(ch['duration_sec'])} |"
            )

        md_lines.append("\n## 二、詳細投影片分段時間表\n")
        md_lines.append("| 頁碼 | 開始時間 (標準) | 開始時間 (幾分幾秒) | 結束時間 | 該頁時長 | 大綱標題 | 分段副標題 |")
        md_lines.append("|:---:|:---:|:---:|:---:|:---:|:---|:---|")
        for item in timeline:
            md_lines.append(
                f"| {item['page_1based']} | `{format_seconds_to_hms(item['start_sec'])}` | `{format_seconds_to_chinese_time(item['start_sec'])}` | `{format_seconds_to_hms(item['end_sec'])}` | {format_seconds_to_chinese_time(item['duration_sec'])} | {item['title'] or '-'} | {item['subtitle'] or '-'} |"
            )

        os.makedirs(os.path.dirname(os.path.abspath(md_path)), exist_ok=True)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write("\n".join(md_lines))

    # CSV 審查表格版 (含 UTF-8 BOM，相容 Excel)
    if csv_path:
        os.makedirs(os.path.dirname(os.path.abspath(csv_path)), exist_ok=True)
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow([
                "項次",
                "投影片頁碼",
                "大綱項目(大標題)",
                "分段內容(副標題)",
                "開始時間(HH:MM:SS)",
                "開始時間(幾分幾秒)",
                "結束時間(HH:MM:SS)",
                "時長(秒)",
                "時長(分秒)",
            ])
            for idx, item in enumerate(timeline, start=1):
                writer.writerow([
                    idx,
                    item["page_1based"],
                    item["title"],
                    item["subtitle"],
                    format_seconds_to_hms(item["start_sec"]),
                    format_seconds_to_chinese_time(item["start_sec"]),
                    format_seconds_to_hms(item["end_sec"]),
                    f"{item['duration_sec']:.2f}",
                    format_seconds_to_chinese_time(item["duration_sec"]),
                ])


# ══════════════════════════════════════════════════════════════
# 一鍵匯出所有交付物（完全對齊「繳交說明.docx」）
# ══════════════════════════════════════════════════════════════

def export_all_deliverables(
    clips: list[str],
    pptx_path: str,
    output_dir: str,
    settings: dict,
    base_name: str = "Azure教學影片",
) -> dict[str, str]:
    """
    一鍵匯出資通院標準驗收檔案包：
      1. 字幕檔：{base_name}.srt
      2. 章節檔：{base_name}.txt (FFMETADATA1 格式)
      3. 簡報檔：{base_name}.pptx (複製/對齊命名)
      4. 完整講稿與 Excel 審查表格 (.csv, .md, .txt)
    """
    from core.parser import parse_all_slides
    from core.script import build_all_scripts

    os.makedirs(output_dir, exist_ok=True)

    print(f"\n[交付物匯出] 1/4 正在讀取 PPT 結構與講稿...")
    parsed_list = parse_all_slides(pptx_path)
    scripts = build_all_scripts(parsed_list, settings)

    ffmpeg_path = settings.get("ffmpeg_path", "ffmpeg")
    print(f"[交付物匯出] 2/4 正在計算 {len(clips)} 個 clip 的影片時長與時間軸...")
    durations = extract_all_clip_durations(clips, ffmpeg_path=ffmpeg_path)
    timeline = build_timeline_metadata(clips, parsed_list, scripts, durations)

    if not timeline:
        print("[警告] 時間軸資料為空，未產生交付檔案。")
        return {}

    total_sec = timeline[-1]["end_sec"]
    print(f"  -> 影片總時長：{format_seconds_to_hms(total_sec)} ({format_seconds_to_chinese_time(total_sec)})")

    # 檔案路徑定義（嚴格符合資通院命名）
    srt_file = os.path.join(output_dir, f"{base_name}.srt")
    vtt_file = os.path.join(output_dir, f"{base_name}.vtt")
    chapter_ffmetadata = os.path.join(output_dir, f"{base_name}.txt")  # 資通院標準章節檔
    pptx_target = os.path.join(output_dir, f"{base_name}.pptx")

    # 額外輔助審查檔案
    transcript_txt = os.path.join(output_dir, f"{base_name}_完整講稿內容.txt")
    transcript_md = os.path.join(output_dir, f"{base_name}_完整講稿內容.md")
    outline_csv = os.path.join(output_dir, f"{base_name}_分段大綱審查表.csv")
    outline_txt = os.path.join(output_dir, f"{base_name}_分段大綱易讀清單.txt")
    outline_md = os.path.join(output_dir, f"{base_name}_分段大綱詳細記錄.md")

    print(f"[交付物匯出] 3/4 正在生成字幕檔 (SRT) 與 FFMETADATA1 章節檔 (.txt)...")
    export_subtitles(timeline, srt_file, vtt_file)
    export_ffmetadata_chapters(timeline, chapter_ffmetadata)

    print(f"[交付物匯出] 4/4 正在匯出完整講稿與 Excel 大綱審查表...")
    export_full_transcript(timeline, transcript_txt, transcript_md)
    export_chapter_outline(timeline, outline_txt, outline_md, outline_csv)

    # 複製 PPTX 到 deliverables 目錄以完成完整交付包
    if os.path.exists(pptx_path) and os.path.abspath(pptx_path) != os.path.abspath(pptx_target):
        shutil.copy2(pptx_path, pptx_target)

    print(f"\n[完成] 資通院標準交付檔案包匯出成功！清單：")
    print(f"  ★ 1. 字幕檔 (SRT)       : {srt_file} (對應「字幕範例.srt」)")
    print(f"  ★ 2. 章節檔 (TXT)       : {chapter_ffmetadata} (對應「大綱_MP4章節範例.txt」，FFMETADATA1 格式)")
    print(f"  ★ 3. 教學簡報 (PPTX)    : {pptx_target}")
    print(f"  - 完整講稿內容 (TXT)   : {transcript_txt}")
    print(f"  - 分段大綱審查表 (CSV) : {outline_csv}")
    print(f"  - 易讀分段大綱 (TXT)   : {outline_txt}")

    return {
        "srt": srt_file,
        "chapter_txt": chapter_ffmetadata,
        "pptx": pptx_target,
        "vtt": vtt_file,
        "transcript_txt": transcript_txt,
        "transcript_md": transcript_md,
        "outline_csv": outline_csv,
        "outline_txt": outline_txt,
    }
