"""Inspect test_line_heights.docx page by page via COM.

Goal: identify what's on p.4 and p.5 (the regressed pages from Round 18
LM0 fix attempt). Extract:
  - Per paragraph: page, y, x, font, size, line spacing rule, runs
  - Per paragraph: dy to next (= LH used by Word)
  - Group by page

Compare to Oxi's prediction (formula vs LM0 lookup) to identify the
quirk that LM0 over-prediction was compensating for.
"""
import json
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = os.path.abspath("tools/golden-test/documents/docx/test_line_heights.docx")
OUT_PATH = os.path.abspath("pipeline_data/test_line_heights_com.json")


def main() -> None:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(DOCX, ReadOnly=True)
        time.sleep(0.5)
        ps = doc.PageSetup
        n_paras = doc.Paragraphs.Count
        print(f"=== test_line_heights.docx ===")
        print(f"  page: {ps.PageWidth:.0f}x{ps.PageHeight:.0f}pt, margins L/T/R/B = {ps.LeftMargin:.0f}/{ps.TopMargin:.0f}/{ps.RightMargin:.0f}/{ps.BottomMargin:.0f}")
        print(f"  paragraphs: {n_paras}")
        print()

        # Use Range.Information(1) = wdActiveEndPageNumber
        # Use Range.Information(3) = wdActiveEndAdjustedPageNumber
        # Use Range.Information(6) = wdVerticalPositionRelativeToPage
        # Use Range.Information(5) = wdHorizontalPositionRelativeToPage
        rows = []
        for pi in range(1, n_paras + 1):
            p = doc.Paragraphs(pi)
            rng = p.Range
            try:
                y = rng.Information(6)
                x = rng.Information(5)
                page = rng.Information(3)  # adjusted page number
            except Exception:
                y = x = page = None
            text = (rng.Text or "").replace("\r", "").replace("\x07", "")
            text = text[:60]

            # Get font/size from first run
            font_name = ""
            font_size = None
            try:
                if rng.Runs.Count > 0:
                    r0 = rng.Runs(1)
                    font_name = r0.Font.Name
                    font_size = r0.Font.Size
            except Exception:
                pass

            # Get paragraph format
            try:
                pf = p.Format
                line_spacing = pf.LineSpacing
                line_spacing_rule = pf.LineSpacingRule  # wdLineSpaceSingle=0, Multiple=5, etc.
                space_before = pf.SpaceBefore
                space_after = pf.SpaceAfter
            except Exception:
                line_spacing = line_spacing_rule = space_before = space_after = None

            rows.append({
                "i": pi,
                "page": page,
                "y_pt": y,
                "x_pt": x,
                "font": font_name,
                "size_pt": font_size,
                "line_spacing": line_spacing,
                "line_spacing_rule": line_spacing_rule,
                "space_before": space_before,
                "space_after": space_after,
                "text": text,
            })

        # Compute dy to next paragraph (only valid within same page)
        for i, r in enumerate(rows):
            if i + 1 < len(rows):
                nx = rows[i + 1]
                if r["y_pt"] is not None and nx["y_pt"] is not None and r["page"] == nx["page"]:
                    r["dy_to_next"] = round(nx["y_pt"] - r["y_pt"], 3)
                else:
                    r["dy_to_next"] = None
            else:
                r["dy_to_next"] = None

        doc.Close(SaveChanges=False)

        # Print page by page
        by_page: dict[int, list] = {}
        for r in rows:
            if r["page"] is not None:
                by_page.setdefault(r["page"], []).append(r)

        for pn in sorted(by_page.keys()):
            print(f"--- Page {pn} ({len(by_page[pn])} paragraphs) ---")
            for r in by_page[pn]:
                print(f"  P{r['i']:3} y={r['y_pt']:6.2f} font={r['font']:<20} sz={r['size_pt']} lsRule={r['line_spacing_rule']} ls={r['line_spacing']} dy_next={r['dy_to_next']}")
                print(f"      text: {r['text']!r}")
            print()

        os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
        with open(OUT_PATH, "w", encoding="utf-8") as f:
            json.dump({"paragraphs": rows, "by_page": {str(k): v for k, v in by_page.items()}}, f, ensure_ascii=False, indent=2)
        print(f"Wrote {OUT_PATH}")
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
