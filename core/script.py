"""
core/script.py
──────────────
講稿組合：將 parse_slide 結果依 settings 規則組合成 TTS 文字
支援 SSML <break> 標籤在每個段落後插入停頓
"""

import re
from core.utils import clean_text_for_tts, extract_title_number


# ══════════════════════════════════════════════════════════════
# settings 輔助
# ══════════════════════════════════════════════════════════════

def get_subtitle_config(subtitle: str, settings: dict) -> dict:
    """取得副標題對應設定，找不到則回傳 default"""
    configs = settings.get("subtitle_config", {})
    return configs.get(subtitle, configs.get("default", {
        "opener": "", "closer": "",
        "read_shapes": False, "read_table": False,
    }))


def _parse_notes_indices(settings: dict) -> set:
    """將 settings 的 notes_slides 解析為 0-based index 集合"""
    cfg = settings.get("notes_slides", None)
    if not cfg:
        return set()
    if isinstance(cfg, list):
        return {int(n) - 1 for n in cfg}
    if isinstance(cfg, str) and "-" in str(cfg):
        start, end = str(cfg).split("-")
        return set(range(int(start) - 1, int(end)))
    return set()


# ══════════════════════════════════════════════════════════════
# SSML 組合
# ══════════════════════════════════════════════════════════════

BREAK_TAG = '<break time="300ms"/>'


def _wrap_ssml(text: str) -> str:
    """將純文字包裝成 SSML 格式，供 Edge-TTS 使用"""
    return (
        '<speak version="1.0" '
        'xmlns="http://www.w3.org/2001/10/synthesis" '
        'xml:lang="zh-TW">'
        + text +
        '</speak>'
    )


def _build_ssml_parts(parts: list, break_ms: int = 300) -> str:
    """
    將 parts 串接並在每個部分後加入 <break>，最後包成 SSML。
    parts 為純文字 list，每個元素代表一個語意段落。
    """
    break_tag = f'<break time="{break_ms}ms"/>'
    joined = break_tag.join(p for p in parts if p.strip())
    return _wrap_ssml(joined)


# ══════════════════════════════════════════════════════════════
# 表格轉語句
# ══════════════════════════════════════════════════════════════

def build_table_sentence(table_rows: list, settings: dict) -> str:
    """將表格轉換成自然語句"""
    if not table_rows:
        return ""

    supported_symbols = set(settings.get(
        "table_supported_symbols", ["☑", "✓", "✔", "V", "v", "O", "o"]
    ))
    prefix = settings.get("table_prefix", "此原則支援以下設定方法：")

    supported   = []
    unsupported = []

    for row in table_rows[1:]:   # 跳過表頭
        if len(row) < 2:
            continue
        method = row[0].strip()
        symbol = row[1].strip()
        if symbol in supported_symbols:
            supported.append(method)
        else:
            unsupported.append(method)

    parts = []
    if supported:
        parts.append(prefix + "、".join(supported) + "。")
    if unsupported:
        parts.append("不支援的設定方法為：" + "、".join(unsupported) + "。")
    return "".join(parts)


# ══════════════════════════════════════════════════════════════
# 單頁講稿組合
# ══════════════════════════════════════════════════════════════

def build_script(
    parsed: dict,
    is_first_page: bool,
    is_continuation: bool,
    is_last_page_of_subtitle: bool,
    settings: dict,
    read_notes: bool = False,
) -> str:
    """
    組合單頁完整講稿（SSML 格式）。

    read_notes=True → 只唸備註欄，忽略投影片內容（適用前導頁）
    """
    break_ms = settings.get("pause_between_paragraphs_ms", 300)

    # ── 備註模式 ─────────────────────────────────────────────
    if read_notes:
        notes = parsed.get("notes", "").strip()
        if not notes:
            return " "
        clean = clean_text_for_tts(notes)
        return _wrap_ssml(clean) if clean else " "

    subtitle = parsed["subtitle"]
    cfg = get_subtitle_config(subtitle, settings)
    continuation_opener = settings.get("continuation_opener", "接續上頁，")

    parts = []

    # ── 大標題（只在單元第一頁唸）───────────────────────────
    if is_first_page and parsed["title"]:
        parts.append(clean_text_for_tts(parsed["title"]) + "。")

    # ── 開場白 ───────────────────────────────────────────────
    if is_continuation:
        parts.append(continuation_opener)
    elif cfg.get("opener"):
        parts.append(cfg["opener"])

    # ── 內容段落（每個 paragraph 獨立加入，各自有停頓）──────
    for text, is_code in parsed.get("paragraphs", []):
        if is_code:
            continue   # 跳過程式碼行，不唸
        clean = clean_text_for_tts(text)
        if clean:
            parts.append(clean)

    # ── 圖形內文字 ───────────────────────────────────────────
    if cfg.get("read_shapes"):
        for shape_text in parsed.get("shapes_text", []):
            clean = clean_text_for_tts(shape_text)
            if clean:
                parts.append(clean)

    # ── 表格 ─────────────────────────────────────────────────
    if cfg.get("read_table") and parsed.get("table_rows"):
        sentence = build_table_sentence(parsed["table_rows"], settings)
        if sentence:
            parts.append(sentence)

    # ── 收尾 ─────────────────────────────────────────────────
    if is_last_page_of_subtitle and cfg.get("closer"):
        parts.append(cfg["closer"])

    if not parts:
        return " "

    return _build_ssml_parts(parts, break_ms)


# ══════════════════════════════════════════════════════════════
# 全頁講稿批次組合
# ══════════════════════════════════════════════════════════════

def build_all_scripts(parsed_list: list, settings: dict) -> list:
    """
    遍歷所有投影片，組合每頁講稿（SSML 格式）。
    回傳 list of str，長度 == 投影片頁數。
    """
    notes_indices = _parse_notes_indices(settings)

    # 預先計算每個副標題的最後一頁 index
    subtitle_last_idx = {}
    for idx, p in enumerate(parsed_list):
        subtitle_last_idx[(p["title"], p["subtitle"])] = idx

    scripts     = []
    prev_title    = None
    prev_subtitle = None

    for idx, parsed in enumerate(parsed_list):
        title    = parsed["title"]
        subtitle = parsed["subtitle"]

        is_first        = (title != prev_title)
        is_continuation = (not is_first and subtitle == prev_subtitle)
        is_last         = (subtitle_last_idx[(title, subtitle)] == idx)
        use_notes       = (idx in notes_indices)

        script = build_script(
            parsed,
            is_first_page=is_first,
            is_continuation=is_continuation,
            is_last_page_of_subtitle=is_last,
            settings=settings,
            read_notes=use_notes,
        )
        scripts.append(script)

        prev_title    = title
        prev_subtitle = subtitle

    return scripts
