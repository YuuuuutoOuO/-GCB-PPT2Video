"""
core/script.py
──────────────
講稿組合：將 parse_slide 結果依 settings 規則組合成段落列表
每頁回傳 list[str]，每個元素為一個獨立段落（純文字，無 SSML）
停頓由 tts.py 在段落間插入靜音音檔實現
"""

import re
from core.utils import clean_text_for_tts


# ══════════════════════════════════════════════════════════════
# settings 輔助
# ══════════════════════════════════════════════════════════════

def _normalize_subtitle(subtitle: str) -> str:
    """
    正規化副標題，去除結尾的頁碼後綴，方便比較與 match。
    例：
      "PowerShell設定方法(2/3)" -> "PowerShell設定方法"
      "說明 (1/2)"              -> "說明"
      "說明"                    -> "說明"
    支援全形、半形括號與空格。
    """
    return re.sub(r"[\s（(]\d+[/／]\d+[)）]?\s*$", "", subtitle).strip()


def get_subtitle_config(subtitle: str, settings: dict) -> dict:
    """
    取得副標題對應設定。
    比對順序：完全匹配 -> 正規化後前綴匹配 -> default
    """
    configs = settings.get("subtitle_config", {})
    default = configs.get("default", {
        "opener": "", "closer": "",
        "read_shapes": False, "read_table": False,
    })

    # 1. 完全匹配
    if subtitle in configs:
        return configs[subtitle]

    # 2. 正規化後前綴匹配（排除 default key）
    normalized = _normalize_subtitle(subtitle)
    for key, cfg in configs.items():
        if key == "default":
            continue
        if normalized == key or normalized.startswith(key):
            return cfg

    return default


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
# 段落過濾輔助
# ══════════════════════════════════════════════════════════════

def _is_valid_paragraph(text: str) -> bool:
    """
    必須含有至少一個中文字、英文字母或數字才算有效段落。
    過濾純符號行（如單獨的 } 、{ 等）。
    """
    text = text.strip()
    if not text or len(text) < 2:
        return False
    if not re.search(r'[\u4e00-\u9fff\w]', text):
        return False
    return True


def _has_chinese(text: str) -> bool:
    """判斷文字是否含有中文字"""
    return bool(re.search(r'[\u4e00-\u9fff]', text))


def _process_paragraphs(paragraphs: list) -> list[str]:
    """
    處理段落列表，實作程式碼區塊智慧跳過邏輯：

    - is_code=True  → 開始跳過，並繼續往後合併
    - 後續段落若無中文 → 視為程式碼區塊的一部分，繼續跳過
    - 後續段落含有中文 → 程式碼區塊結束，恢復正常唸出
    - is_code=False → 正常唸出

    範例（PowerShell 跨行指令）：
      "# az network watcher list --query []."  is_code=True  → 跳過
      "{Location:location,State:provisioningState}" -o table  is_code=False, 無中文 → 跳過
      "確保每個 Network Watcher 的 provisioningState..."      is_code=False, 有中文 → 唸出 ✅
    """
    result = []
    i = 0
    while i < len(paragraphs):
        text, is_code = paragraphs[i]

        if is_code:
            # 進入程式碼區塊，往後掃描直到遇到含中文的段落
            i += 1
            while i < len(paragraphs):
                next_text, next_is_code = paragraphs[i]
                # 含中文 → 程式碼區塊結束，讓這行正常處理
                if _has_chinese(next_text):
                    break
                # 不含中文（含下一行 is_code 或純英數符號）→ 繼續跳過
                i += 1
            continue

        # 正常段落
        clean = clean_text_for_tts(text)
        if _is_valid_paragraph(clean):
            result.append(clean)
        i += 1

    return result


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
) -> list[str]:
    """
    組合單頁講稿，回傳段落列表（純文字）。
    每個元素代表一個語意段落，tts.py 會在段落間插入靜音停頓。

    read_notes=True → 只唸備註欄，回傳單一元素的列表
    空頁             → 回傳空列表
    """
    # ── 備註模式 ─────────────────────────────────────────────
    if read_notes:
        notes = parsed.get("notes", "").strip()
        if not notes:
            return []
        clean = clean_text_for_tts(notes)
        return [clean] if clean else []

    subtitle = parsed["subtitle"]
    cfg = get_subtitle_config(subtitle, settings)
    continuation_opener = settings.get("continuation_opener", "接續上頁，") or ""

    def _add(parts: list, text: str):
        """統一過濾 None、空字串、無意義符號"""
        if text and _is_valid_paragraph(text):
            parts.append(text.strip())

    parts = []

    # ── 大標題（只在單元第一頁唸）───────────────────────────
    if is_first_page and parsed["title"]:
        _add(parts, clean_text_for_tts(parsed["title"]) + "。")

    # ── 開場白 ───────────────────────────────────────────────
    if is_continuation:
        _add(parts, continuation_opener)
    else:
        _add(parts, cfg.get("opener") or "")

    # ── 內容段落（程式碼區塊智慧跳過）───────────────────────
    processed = _process_paragraphs(parsed.get("paragraphs", []))
    for text in processed:
        _add(parts, text)

    # ── 圖形內文字 ───────────────────────────────────────────
    if cfg.get("read_shapes"):
        for shape_text in parsed.get("shapes_text", []):
            _add(parts, clean_text_for_tts(shape_text))

    # ── 表格 ─────────────────────────────────────────────────
    if cfg.get("read_table") and parsed.get("table_rows"):
        _add(parts, build_table_sentence(parsed["table_rows"], settings))

    # ── 收尾 ─────────────────────────────────────────────────
    if is_last_page_of_subtitle:
        _add(parts, cfg.get("closer") or "")

    return parts


# ══════════════════════════════════════════════════════════════
# 全頁講稿批次組合
# ══════════════════════════════════════════════════════════════

def build_all_scripts(parsed_list: list, settings: dict) -> list[list[str]]:
    """
    遍歷所有投影片，組合每頁講稿。
    回傳 list[list[str]]，外層長度 == 投影片頁數，內層為該頁的段落列表。
    """
    notes_indices = _parse_notes_indices(settings)

    # 預先計算每個（標題, 正規化副標題）的最後一頁 index
    subtitle_last_idx = {}
    for idx, p in enumerate(parsed_list):
        key = (p["title"], _normalize_subtitle(p["subtitle"]))
        subtitle_last_idx[key] = idx

    scripts       = []
    prev_title    = None
    prev_norm_sub = None   # 上一頁正規化後的副標題

    for idx, parsed in enumerate(parsed_list):
        title    = parsed["title"]
        subtitle = parsed["subtitle"]
        norm_sub = _normalize_subtitle(subtitle)

        is_first        = (title != prev_title)
        is_continuation = (not is_first and norm_sub == prev_norm_sub)
        is_last         = (subtitle_last_idx[(title, norm_sub)] == idx)
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
        prev_norm_sub = norm_sub

    return scripts