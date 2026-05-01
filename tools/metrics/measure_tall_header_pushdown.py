"""§8.2 Tall-header pushdown — body Y position when header content overflows
the gap between headerDistance and topMargin.

Spec line 936 says: "Earlier note ('3-line 14pt header → body_y=90pt when
topMargin=72') was measured for noGrid; the formula has not been re-verified
under the corrected spec and remains a candidate for follow-up Ra2 measurement."

This script sweeps:
  - Header line count: 1..5
  - Header font/size: Calibri 11pt, MS Mincho 10.5pt, 14pt, 18pt
  - topMargin: 36, 72, 108pt
  - headerDistance: 18, 36, 54pt
  - Grid: noGrid vs LayoutMode=1 (linesAndChars) pitch=18pt

For each combo, measures:
  - header_top_y (Information(6) of header.Range.Paragraphs(1))
  - header_bot_y (Information(6) of last header paragraph)
  - body_y (Information(6) of body Paragraph 1)

Then derives the relationship: body_y = max(topMargin, header_bot + something).

Output: pipeline_data/ra_manual_measurements.json key
"tall_header_pushdown_2026-05-02"
"""
import win32com.client
import os
import time
import json
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Configurations
TOP_MARGINS = [36, 72, 108]
HEADER_DISTS = [18, 36, 54]
HEADER_LINE_COUNTS = [1, 2, 3, 4]
HEADER_FONT_SIZES = [
    ("Calibri", 11),
    ("ＭＳ 明朝", 10.5),
    ("Calibri", 14),
    ("Calibri", 18),
]
GRID_MODES = ["noGrid", "linesAndChars_18"]

RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")


def make_word():
    w = win32com.client.Dispatch("Word.Application")
    w.Visible = False
    w.DisplayAlerts = False
    return w


def measure_one(word, top_margin, hdr_dist, n_lines, font, size, grid_mode):
    doc = word.Documents.Add()
    try:
        sec = doc.Sections(1)
        ps = sec.PageSetup
        ps.PageHeight = 841.9
        ps.PageWidth = 595.3
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = top_margin
        ps.BottomMargin = 72
        ps.HeaderDistance = hdr_dist
        ps.FooterDistance = 36
        if grid_mode == "linesAndChars_18":
            ps.LayoutMode = 2
            ps.LinesPage = 41  # forces ~ pitch=18pt with topMargin=72
        else:
            ps.LayoutMode = 0  # default

        hdr = sec.Headers(1)
        hdr_text = "\r".join(f"H{i + 1}" for i in range(n_lines))
        hdr.Range.Text = hdr_text
        hdr.Range.Font.Name = font
        hdr.Range.Font.Size = size
        for hp_i in range(1, hdr.Range.Paragraphs.Count + 1):
            hp = hdr.Range.Paragraphs(hp_i)
            hp.Format.SpaceBefore = 0
            hp.Format.SpaceAfter = 0
            hp.Format.LineSpacingRule = 0  # wdLineSpaceSingle

        doc.Content.Text = "Body"
        bp = doc.Paragraphs(1)
        bp.Range.Font.Name = "Calibri"
        bp.Range.Font.Size = 11
        bp.Format.SpaceBefore = 0
        bp.Format.SpaceAfter = 0

        doc.Repaginate()
        time.sleep(0.10)

        # Measure header paragraphs
        hdr_paras = []
        for i in range(1, hdr.Range.Paragraphs.Count + 1):
            p = hdr.Range.Paragraphs(i)
            try:
                hdr_paras.append({
                    "i": i,
                    "y": round(float(p.Range.Information(6)), 4),
                })
            except Exception as e:
                hdr_paras.append({"i": i, "err": str(e)})

        body_y = round(float(bp.Range.Information(6)), 4)
        # Estimate header_height from last_y - first_y + estimated line height
        if hdr_paras and "y" in hdr_paras[0] and "y" in hdr_paras[-1]:
            first_y = hdr_paras[0]["y"]
            last_y = hdr_paras[-1]["y"]
        else:
            first_y = last_y = None
        return {
            "n_lines": n_lines,
            "font": font,
            "size": size,
            "topMargin": top_margin,
            "headerDistance": hdr_dist,
            "grid": grid_mode,
            "header_paras": hdr_paras,
            "header_first_y": first_y,
            "header_last_y": last_y,
            "body_y": body_y,
            "body_minus_topMargin": round(body_y - top_margin, 4),
        }
    finally:
        doc.Close(SaveChanges=False)


def main():
    results = []
    for grid_mode in GRID_MODES:
        word = make_word()
        try:
            for font, size in HEADER_FONT_SIZES:
                for top_margin in TOP_MARGINS:
                    for hdr_dist in HEADER_DISTS:
                        for n_lines in HEADER_LINE_COUNTS:
                            try:
                                r = measure_one(word, top_margin, hdr_dist,
                                                n_lines, font, size, grid_mode)
                                results.append(r)
                                line = (f"[{grid_mode}][{font}/{size}]"
                                        f"[tm={top_margin} hd={hdr_dist} "
                                        f"n={n_lines}] body_y={r['body_y']} "
                                        f"(Δ={r['body_minus_topMargin']}) "
                                        f"hdr_first={r['header_first_y']} "
                                        f"last={r['header_last_y']}")
                                print(line, flush=True)
                            except Exception as e:
                                err = {
                                    "grid": grid_mode,
                                    "font": font,
                                    "size": size,
                                    "topMargin": top_margin,
                                    "headerDistance": hdr_dist,
                                    "n_lines": n_lines,
                                    "error": str(e),
                                }
                                results.append(err)
                                print(f"  ERROR {err}", flush=True)
        finally:
            try:
                word.Quit()
            except Exception:
                pass
            time.sleep(1.0)

    if os.path.exists(RESULT_PATH):
        try:
            with open(RESULT_PATH, encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            existing = {}
    else:
        existing = {}
    existing["tall_header_pushdown_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {len(results)} records to {RESULT_PATH}")


if __name__ == "__main__":
    main()
