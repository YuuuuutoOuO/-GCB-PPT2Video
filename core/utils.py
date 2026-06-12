"""
core/utils.py
─────────────
共用工具：文字清理、範圍參數解析、Registry DPI 設定
"""

import os
import re
import sys


# ══════════════════════════════════════════════════════════════
# 文字清理
# ══════════════════════════════════════════════════════════════

def clean_text_for_tts(text: str) -> str:
    """
    清理不適合 TTS 朗讀的符號：
      【】  → 移除
      /     → 替換為頓號
      換行  → 替換為頓號
      連續標點 → 合併為單一頓號
      多餘空白 → 壓縮為單一空格
    """
    text = text.replace("【", "").replace("】", "")
    text = text.replace("/", "，")
    text = text.replace("\r\n", "，").replace("\n", "，").replace("\r", "，")
    text = re.sub(r"[，。、]{2,}", "，", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


# ══════════════════════════════════════════════════════════════
# 範圍參數解析
# ══════════════════════════════════════════════════════════════

def parse_range_str(range_str: str) -> set:
    """
    將 "1,5-10,100-200" 解析為整數集合。
    例：{1, 5, 6, 7, 8, 9, 10, 100, ..., 200}
    """
    result = set()
    for part in range_str.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            result.update(range(int(start), int(end) + 1))
        elif part.isdigit():
            result.add(int(part))
    return result


def extract_title_number(title: str) -> int:
    """
    從標題開頭抓數字。
    例："58.【PostgreSQL】..." → 58
    無數字回傳 -1（代表前導頁）
    """
    m = re.match(r"^(\d+)", title.strip())
    return int(m.group(1)) if m else -1


def resolve_target_indices(parsed_list: list, pages_str: str, range_str: str) -> set:
    """
    根據 --pages 和 --range 參數，計算出目標頁的 0-based index 集合。
    回傳空集合代表全部處理。
    """
    if not pages_str and not range_str:
        return set()

    target = set()

    if pages_str:
        for pn in parse_range_str(pages_str):
            idx = pn - 1
            if 0 <= idx < len(parsed_list):
                target.add(idx)

    if range_str:
        title_nums = parse_range_str(range_str)
        for idx, parsed in enumerate(parsed_list):
            num = extract_title_number(parsed["title"])
            if num in title_nums:
                target.add(idx)

    return target


def confirm_dual_mode(parsed_list: list, pages_str: str, range_str: str, target_indices: set):
    """
    同時使用 --pages 和 --range 時顯示警告並要求 y 確認。
    """
    page_indices = set()
    if pages_str:
        for pn in parse_range_str(pages_str):
            idx = pn - 1
            if 0 <= idx < len(parsed_list):
                page_indices.add(idx)

    range_indices = set()
    if range_str:
        title_nums = parse_range_str(range_str)
        for idx, parsed in enumerate(parsed_list):
            num = extract_title_number(parsed["title"])
            if num in title_nums:
                range_indices.add(idx)

    page_display  = ", ".join(str(i + 1) for i in sorted(page_indices))  or "（無）"
    range_display = ", ".join(str(i + 1) for i in sorted(range_indices)) or "（無）"

    print("=" * 60)
    print("⚠️  警告：同時指定了 --pages 和 --range，將取兩者聯集執行。")
    print(f"   --pages 涵蓋頁碼：第 {page_display} 頁")
    print(f"   --range 涵蓋頁碼：第 {range_display} 頁")
    print(f"   聯集共計處理  ：{len(target_indices)} 頁")
    print("=" * 60)
    ans = input("確認執行？[y/N]：").strip().lower()
    if ans != "y":
        print("已取消執行。")
        sys.exit(0)


# ══════════════════════════════════════════════════════════════
# Windows Registry DPI 設定
# ══════════════════════════════════════════════════════════════

def set_pptx_export_dpi(dpi: int = 200):
    """
    透過 Windows Registry 調整 PowerPoint 匯出圖片的 DPI。
    預設值為 96，調高後截圖畫質明顯提升。
    需要系統管理員權限執行。

    Registry 路徑：
      HKCU\\Software\\Microsoft\\Office\\<version>\\PowerPoint\\Options
      值名稱：ExportBitmapResolution (DWORD)
    """
    if sys.platform != "win32":
        print("  [略過] DPI 設定僅支援 Windows 系統")
        return

    try:
        import winreg

        # 自動偵測已安裝的 Office 版本（嘗試常見版本號）
        office_versions = ["16.0", "15.0", "14.0", "12.0"]
        key_path = None

        for ver in office_versions:
            path = rf"Software\Microsoft\Office\{ver}\PowerPoint\Options"
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path,
                                     0, winreg.KEY_READ)
                winreg.CloseKey(key)
                key_path = path
                break
            except FileNotFoundError:
                continue

        if not key_path:
            # 找不到已安裝版本，嘗試建立 16.0（最常見）
            key_path = r"Software\Microsoft\Office\16.0\PowerPoint\Options"
            print(f"  [警告] 找不到已安裝的 Office 版本，嘗試建立 {key_path}")

        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path,
                             0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, "ExportBitmapResolution", 0, winreg.REG_DWORD, dpi)
        winreg.CloseKey(key)
        print(f"  ✅ 匯出 DPI 已設定為 {dpi}")

    except PermissionError:
        print("  ❌ 權限不足，請以系統管理員身份執行程式")
    except Exception as e:
        print(f"  ❌ Registry 設定失敗：{e}")
