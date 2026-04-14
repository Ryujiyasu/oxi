"""Measure character positions in db9ca18368cd to diagnose line break differences.

Word fits 112 chars on P4 L1, Oxi only 94. Need to measure:
1. Actual content width
2. Per-character advance widths (via Information(5))
3. Line break positions
"""
import win32com.client
import os
import json
import time

DOCX = os.path.join(os.path.dirname(__file__), "..", "golden-test", "documents", "docx",
                     "db9ca18368cd_20241122_resource_open_data_01.docx")

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    time.sleep(1)

    results = {}

    # 1. Page setup
    sec = doc.Sections(1)
    ps = sec.PageSetup
    results["page_setup"] = {
        "page_width_pt": ps.PageWidth,
        "page_height_pt": ps.PageHeight,
        "margin_left_pt": ps.LeftMargin,
        "margin_right_pt": ps.RightMargin,
        "margin_top_pt": ps.TopMargin,
        "margin_bottom_pt": ps.BottomMargin,
        "content_width_pt": ps.PageWidth - ps.LeftMargin - ps.RightMargin,
    }
    print(f"Page: {ps.PageWidth}x{ps.PageHeight}, margins L={ps.LeftMargin} R={ps.RightMargin}")
    print(f"Content width: {ps.PageWidth - ps.LeftMargin - ps.RightMargin}")

    # 2. docDefaults font info
    try:
        default_style = doc.Styles(-1)  # wdStyleNormal
        results["normal_style"] = {
            "font_name": default_style.Font.Name,
            "font_size": default_style.Font.Size,
        }
        print(f"Normal style: {default_style.Font.Name} {default_style.Font.Size}pt")
    except Exception as e:
        print(f"Normal style error: {e}")
        results["normal_style"] = {"error": str(e)}

    # 3. Measure first 30 paragraphs: Y position, font, size, line count
    para_data = []
    for i in range(1, min(31, doc.Paragraphs.Count + 1)):
        p = doc.Paragraphs(i)
        r = p.Range
        try:
            y = r.Information(6)  # wdVerticalPositionRelativeToPage
            x = r.Information(5)  # wdHorizontalPositionRelativeToPage
            font_name = r.Font.Name
            font_size = r.Font.Size
            text = r.Text[:80].replace("\r", "\\r").replace("\n", "\\n")
            text_len = len(r.Text.rstrip("\r\n"))
            page = r.Information(3)  # wdActiveEndPageNumber

            pd = {
                "index": i,
                "page": page,
                "y": y,
                "x": x,
                "font": font_name,
                "size": font_size,
                "text_len": text_len,
                "text_preview": text[:60],
            }
            para_data.append(pd)
            print(f"  P{i:2d}: page={page} y={y:.1f} x={x:.1f} {font_name} {font_size}pt len={text_len} \"{text[:40]}\"")
        except Exception as e:
            print(f"  P{i:2d}: error {e}")
            para_data.append({"index": i, "error": str(e)})
    results["paragraphs"] = para_data

    # 4. Measure per-character advance widths for P5 (the URL paragraph)
    # Find the paragraph with long URL text
    target_para = None
    for i in range(1, min(20, doc.Paragraphs.Count + 1)):
        p = doc.Paragraphs(i)
        text = p.Range.Text
        if "https://" in text and len(text) > 100:
            target_para = i
            break

    if target_para:
        p = doc.Paragraphs(target_para)
        text = p.Range.Text.rstrip("\r\n")
        print(f"\n=== Per-char measurement for P{target_para} ({len(text)} chars) ===")

        char_data = []
        start = p.Range.Start
        prev_x = None
        for j in range(min(len(text), 130)):
            rng = doc.Range(start + j, start + j + 1)
            try:
                cx = rng.Information(5)  # horizontal position
                char = text[j] if j < len(text) else "?"
                advance = cx - prev_x if prev_x is not None and j > 0 else 0
                char_data.append({
                    "pos": j,
                    "char": char,
                    "x": cx,
                    "advance": round(advance, 2) if advance != 0 else 0,
                })
                if j < 20 or (j > 90 and j < 120):
                    print(f"  [{j:3d}] '{char}' x={cx:.1f} adv={advance:.2f}")
                prev_x = cx
            except Exception as e:
                print(f"  [{j:3d}] error: {e}")
                break
        results["url_chars"] = char_data

    # 5. Measure a body text paragraph (P8 or similar long paragraph)
    body_para = None
    for i in range(5, min(25, doc.Paragraphs.Count + 1)):
        p = doc.Paragraphs(i)
        text = p.Range.Text
        if len(text) > 200 and "https://" not in text:
            body_para = i
            break

    if body_para:
        p = doc.Paragraphs(body_para)
        text = p.Range.Text.rstrip("\r\n")
        print(f"\n=== Per-char measurement for P{body_para} body text ({len(text)} chars) ===")

        body_chars = []
        start = p.Range.Start
        prev_x = None
        for j in range(min(len(text), 130)):
            rng = doc.Range(start + j, start + j + 1)
            try:
                cx = rng.Information(5)
                char = text[j] if j < len(text) else "?"
                advance = cx - prev_x if prev_x is not None and j > 0 else 0
                body_chars.append({
                    "pos": j,
                    "char": char,
                    "x": cx,
                    "advance": round(advance, 2) if advance != 0 else 0,
                })
                if j < 10 or (j > 105 and j < 120):
                    print(f"  [{j:3d}] '{char}' x={cx:.1f} adv={advance:.2f}")
                prev_x = cx
            except Exception as e:
                print(f"  [{j:3d}] error: {e}")
                break
        results["body_chars"] = body_chars

        # Find the actual line break position
        # Check where the X resets (new line)
        line_breaks = []
        prev_x_val = 0
        for cd in body_chars:
            if cd["x"] < prev_x_val - 10:  # X jumped back = new line
                line_breaks.append(cd["pos"])
            prev_x_val = cd["x"]
        print(f"  Line breaks at positions: {line_breaks}")
        results["body_line_breaks"] = line_breaks

    # Save results
    out_path = os.path.join(os.path.dirname(__file__), "..", "..", "pipeline_data",
                           "ra_manual_measurements_db9c_linebreak.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {out_path}")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
