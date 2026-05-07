"""COM-measure Word's per-char positions for bd90b00's idx=24 paragraph
(the 1-based equivalent: paragraph i=24 = '統計センターがあらかじめ定めるア以外の方法').

Goal: identify per-char width residual that lets Word fit '法' on line 1
while Oxi wraps it to line 2. This 12pt extra height pushes 備考 to p2,
breaking the only Phase 1 mismatch on bd90b00 (score 0.9630 → 1.0).

Tool calls:
- Open the docx via Word COM
- Iterate every char in para i=24
- Use Range.Information(WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE) to get x
- Use Range.Information(WD_VERTICAL_POSITION_RELATIVE_TO_PAGE) to get y
- For each char, also note its actual rendered font and font size

Run: python tools/metrics/measure_bd90b00_para24_chars.py
"""
import os
import sys
import json
import win32com.client

WD_HPOS = 5  # WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE
WD_VPOS = 6  # WD_VERTICAL_POSITION_RELATIVE_TO_PAGE
WD_PAGE_NUM = 3  # WD_ACTIVE_END_PAGE_NUMBER

DOCX = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx", "bd90b00ab7a7_order_05.docx"
))

OUT = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "..",
    "pipeline_data", "bd90b00_para24_word_chars.json"
))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    print(f"Opening: {DOCX}")

    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        doc = word.Documents.Open(DOCX, ReadOnly=True)
        para = doc.Paragraphs(24)  # 1-based; ECMA-376 says para 24
        rng = para.Range
        text = rng.Text
        print(f"Para 24 text ({len(text)} chars): {text!r}")

        # Use start-collapsed range for first-line position (R30 fix)
        first_rng = doc.Range(rng.Start, rng.Start)
        first_y = first_rng.Information(WD_VPOS)
        first_x = first_rng.Information(WD_HPOS)
        first_page = first_rng.Information(WD_PAGE_NUM)
        print(f"Para start: page={first_page} x={first_x} y={first_y}")

        # Per-char positions
        chars_data = []
        for i in range(len(text)):
            ch_rng = doc.Range(rng.Start + i, rng.Start + i + 1)
            ch_text = ch_rng.Text
            try:
                ch_x = ch_rng.Information(WD_HPOS)
                ch_y = ch_rng.Information(WD_VPOS)
                ch_page = ch_rng.Information(WD_PAGE_NUM)
            except Exception as e:
                ch_x = ch_y = ch_page = None
            font_name = ch_rng.Font.Name if ch_rng.Font else None
            font_size = ch_rng.Font.Size if ch_rng.Font else None
            chars_data.append({
                "i": i,
                "char": ch_text,
                "page": ch_page,
                "x": ch_x,
                "y": ch_y,
                "font_name": font_name,
                "font_size": font_size,
            })
            if i < 5 or i > len(text) - 5:
                print(f"  i={i:3d} char={ch_text!r:>4} x={ch_x} y={ch_y} font={font_name} sz={font_size}")

        # Print line-break detection: find chars where y changes
        prev_y = None
        line_starts = []
        for c in chars_data:
            if c["y"] != prev_y:
                line_starts.append(c)
                prev_y = c["y"]
        print(f"\nLine starts ({len(line_starts)}):")
        for ls in line_starts:
            print(f"  i={ls['i']:3d} y={ls['y']} char={ls['char']!r}")

        # Find the LAST char on line 1 (i.e., before y change)
        if len(line_starts) > 1:
            second_line_i = line_starts[1]["i"]
            last_l1 = chars_data[second_line_i - 1]
            print(f"\nLast char on line 1: i={last_l1['i']} char={last_l1['char']!r} x={last_l1['x']}")

        out = {
            "doc": "bd90b00ab7a7_order_05.docx",
            "para_idx": 24,
            "para_text": text,
            "para_start": {"page": first_page, "x": first_x, "y": first_y},
            "chars": chars_data,
            "line_starts": [{"i": ls["i"], "y": ls["y"], "char": ls["char"]} for ls in line_starts],
        }
        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        print(f"\nWrote {OUT}")

        doc.Close(SaveChanges=False)
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
