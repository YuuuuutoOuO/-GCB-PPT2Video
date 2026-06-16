"""
debug_inspect_paragraphs.py
─────────────────────────────
精確診斷：直接呼叫 core.parser.parse_slide()，
印出 paragraphs 的原始內容（含 is_code 標記），
並模擬 script.py 的 _process_paragraphs() 處理過程，
找出內容在哪一步被吞掉。

使用方式：
  python debug_inspect_paragraphs.py --pptx file.pptx --pages 685,686
"""

import argparse
import re
from pptx import Presentation

from core.parser import parse_slide
from core.utils import clean_text_for_tts


def _is_valid_paragraph(text: str) -> bool:
    text = text.strip()
    if not text or len(text) < 2:
        return False
    if not re.search(r'[\u4e00-\u9fff\w]', text):
        return False
    return True


def _has_chinese(text: str) -> bool:
    return bool(re.search(r'[\u4e00-\u9fff]', text))


def _process_paragraphs_verbose(paragraphs):
    """模擬 script.py 的 _process_paragraphs，但印出每一步決策"""
    result = []
    i = 0
    while i < len(paragraphs):
        text, is_code = paragraphs[i]
        print(f"    [{i}] is_code={is_code}  text={text[:80]!r}")

        if is_code:
            print(f"        → 判定為程式碼，開始跳過...")
            i += 1
            while i < len(paragraphs):
                next_text, next_is_code = paragraphs[i]
                if _has_chinese(next_text):
                    print(f"        → [{i}] 含中文，程式碼區塊結束：{next_text[:50]!r}")
                    break
                print(f"        → [{i}] 無中文，繼續跳過：{next_text[:50]!r}")
                i += 1
            continue

        clean = clean_text_for_tts(text)
        valid = _is_valid_paragraph(clean)
        print(f"        → clean_text_for_tts 後：{clean[:80]!r}")
        print(f"        → _is_valid_paragraph：{valid}")
        if valid:
            result.append(clean)
            print(f"        ✅ 加入結果")
        else:
            print(f"        ❌ 被過濾掉")
        i += 1

    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--pages", required=True)
    args = parser.parse_args()

    prs = Presentation(args.pptx)
    pages = [int(p.strip()) for p in args.pages.split(",")]

    for page in pages:
        slide = prs.slides[page - 1]
        parsed = parse_slide(slide)

        print(f"\n{'='*70}")
        print(f"第 {page} 頁")
        print(f"{'='*70}")
        print(f"title    = {parsed['title']!r}")
        print(f"subtitle = {parsed['subtitle']!r}")
        print(f"\nparagraphs 原始內容（共 {len(parsed['paragraphs'])} 個）：")
        for i, (text, is_code) in enumerate(parsed['paragraphs']):
            print(f"  [{i}] is_code={is_code}  text={text!r}")

        print(f"\n模擬 _process_paragraphs 處理過程：")
        final = _process_paragraphs_verbose(parsed['paragraphs'])
        print(f"\n最終結果（共 {len(final)} 段）：")
        for f in final:
            print(f"  - {f!r}")


if __name__ == "__main__":
    main()