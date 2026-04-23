"""Measure d77a p6/p7 cell-split cascade via COM.

Goal: identify why Word moves paragraph(s) to p7 that Oxi keeps on p6,
producing 14.5pt close_border gap.

For each paragraph in the cell that spans the p6/p7 split:
- page number (Information(3))
- Y start (Information(6))
- widowControl flag, keepLines, keepNext
- line count, font size, line_height setting
- number of text chars

Also: for the WHOLE document, find which paragraph is the LAST on p6 per
Word (start_page=6, end_page=6) and the FIRST on p7 (start_page=7). The
transition paragraph is the split point.
"""
import win32com.client
import json
import os
from pathlib import Path

DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
OUT = r"C:\Users\ryuji\oxi-main\pipeline_data\d77a_p6_p7_cascade.json"


def main():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(DOC, ReadOnly=True)
        total = doc.Paragraphs.Count
        print(f"total paragraphs: {total}")

        data = []
        for i in range(1, total + 1):
            p = doc.Paragraphs(i)
            rng = p.Range
            try:
                page_start = rng.Information(3)
                y_start = rng.Information(6)
            except Exception:
                page_start = -1; y_start = -1.0
            try:
                safe_end = max(rng.Start, rng.End - 1)
                end_rng = doc.Range(safe_end, safe_end)
                page_end = end_rng.Information(3)
                y_end = end_rng.Information(6)
            except Exception:
                page_end = -1; y_end = -1.0

            # ParagraphFormat flags
            pf = p.Format
            widow = None
            keep_lines = None
            keep_next = None
            space_before = None
            space_after = None
            line_spacing_rule = None
            line_spacing = None
            try:
                widow = pf.WidowControl  # True/False
                keep_lines = pf.KeepTogether  # keepLines? Actually KeepWithNext vs KeepLines: KeepTogether is in OOXML terms
                keep_next = pf.KeepWithNext
                space_before = pf.SpaceBefore
                space_after = pf.SpaceAfter
                line_spacing_rule = pf.LineSpacingRule
                line_spacing = pf.LineSpacing
            except Exception as e:
                print(f"  para {i}: ParagraphFormat err {e}")

            # Font size of first run
            font_size = None
            try:
                font_size = rng.Font.Size
            except Exception:
                pass

            text = rng.Text[:60].replace("\r", " ").replace("\n", " ").replace("\x07", "|")
            data.append({
                "idx": i,
                "page_start": page_start,
                "page_end": page_end,
                "y_start": round(y_start, 2) if isinstance(y_start, float) else y_start,
                "y_end": round(y_end, 2) if isinstance(y_end, float) else y_end,
                "widow": widow,
                "keep_together": keep_lines,
                "keep_with_next": keep_next,
                "sb": space_before,
                "sa": space_after,
                "line_rule": line_spacing_rule,  # 0=single, 1=1.5, 2=double, 3=atLeast, 4=exactly, 5=multiple
                "line_spacing": line_spacing,
                "font_size": font_size,
                "text": text,
            })

        # Find p6/p7 boundary
        print("\n=== p5-p7 paragraphs ===")
        for d in data:
            if d["page_start"] in (5, 6, 7, 8) or d["page_end"] in (5, 6, 7, 8):
                print(f"  idx={d['idx']:3} pstart={d['page_start']} y={d['y_start']:>7} pend={d['page_end']} y={d['y_end']:>7} "
                      f"widow={d['widow']} kt={d['keep_together']} kn={d['keep_with_next']} "
                      f"fs={d['font_size']} rule={d['line_rule']} ls={d['line_spacing']}")
                print(f"       text: {d['text']!r}")

        doc.Close(SaveChanges=False)
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nSaved {len(data)} paragraphs to {OUT}")


if __name__ == "__main__":
    main()
