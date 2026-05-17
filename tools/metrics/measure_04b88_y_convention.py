"""Measure Word's Y-reporting convention for 04b88e7e0b25 paragraphs.

Hypothesis: Word's Information(6) reports the LINE BOX TOP, while Oxi reports
GLYPH TOP (= line_box_top + text_y_off). For paragraphs with exact line=340tw
(=17pt) and font 10.5pt, text_y_off = 6.5pt → systematic +6.5pt dy.

For each paragraph we capture:
- start_y    = Information(6) at collapsed start range
- char1_y    = Information(6) at a 1-char range from start (glyph-top? baseline?)
- line_y     = Information(7) wdVerticalPositionRelativeToTextBoundary at start
- line_h_set = Word.Format.LineSpacing setting value (advisory only)
- font_size  = paragraph's first run font size
- line_rule  = the spacing rule string ("Exact"/"Multiple"/...)

Compares to Oxi's dump-layout to identify the convention mismatch.

Run from repo root:
    python tools/metrics/measure_04b88_y_convention.py
"""
from __future__ import annotations

import json
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOC_PATH = os.path.join(
    REPO_ROOT, "tools", "golden-test", "documents", "docx",
    "04b88e7e0b25_index-19.docx",
)
OUT_PATH = os.path.join(REPO_ROOT, "tools", "metrics", "04b88_y_convention.json")

# wd* constants we use:
WD_VERT_PAGE = 6        # wdVerticalPositionRelativeToPage
WD_VERT_TEXT = 7        # wdVerticalPositionRelativeToTextBoundary
WD_FIRST_CHAR = 10      # wdFirstCharacterColumnNumber (sanity)
WD_HORIZ_PAGE = 5       # wdHorizontalPositionRelativeToPage


def measure() -> dict:
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    out = {"doc": "04b88e7e0b25", "paragraphs": []}
    doc = word.Documents.Open(DOC_PATH, ReadOnly=True)
    time.sleep(0.3)
    try:
        n_paras = doc.Paragraphs.Count
        # Sample: first 40 (page 1 body) + page-2 table cells
        # Look at paragraphs 50-100 (likely in table on p2-3) for cell behavior
        ranges = list(range(1, min(40, n_paras) + 1)) + list(range(50, min(100, n_paras) + 1))
        for pi in ranges:
            p = doc.Paragraphs(pi)
            rng = p.Range
            raw_text = (rng.Text or "")[:30]
            # Collapsed start range
            start_rng = doc.Range(rng.Start, rng.Start)
            # End-of-first-character range
            ch1_rng = doc.Range(rng.Start, rng.Start + 1) if rng.End > rng.Start else start_rng

            rec = {
                "i": pi,
                "text": raw_text,
                "start": rng.Start,
                "end": rng.End,
            }
            try:
                rec["page_start"] = start_rng.Information(3)
            except Exception:
                rec["page_start"] = None
            try:
                rec["y_page_start"] = start_rng.Information(WD_VERT_PAGE)
            except Exception:
                rec["y_page_start"] = None
            try:
                rec["y_text_start"] = start_rng.Information(WD_VERT_TEXT)
            except Exception:
                rec["y_text_start"] = None
            try:
                rec["y_page_ch1"] = ch1_rng.Information(WD_VERT_PAGE)
            except Exception:
                rec["y_page_ch1"] = None
            try:
                rec["x_page_start"] = start_rng.Information(WD_HORIZ_PAGE)
            except Exception:
                rec["x_page_start"] = None
            # Format settings
            fmt = p.Format
            rec["line_spacing"] = fmt.LineSpacing  # value (pt or factor)
            rec["line_spacing_rule"] = fmt.LineSpacingRule  # 0-5 int
            rec["space_before"] = fmt.SpaceBefore
            rec["space_after"] = fmt.SpaceAfter
            # First run font
            try:
                first_run = p.Range.Characters(1).Font
                rec["fs"] = first_run.Size
                rec["font"] = first_run.NameFarEast or first_run.Name
            except Exception:
                rec["fs"] = None
                rec["font"] = None
            # Paragraph alignment
            rec["align"] = fmt.Alignment
            # Is paragraph in a table?
            try:
                rec["in_table"] = bool(rng.Information(12))  # wdWithInTable
            except Exception:
                rec["in_table"] = None
            out["paragraphs"].append(rec)
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()
    return out


if __name__ == "__main__":
    res = measure()
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=2)
    print(f"Wrote {OUT_PATH} with {len(res['paragraphs'])} paragraphs")
    # Brief summary
    print(f"\n{'i':>3} {'page':>4} {'y_pg':>6} {'y_text':>6} {'y_ch1':>6} {'fs':>5} {'rule':>4} {'spacing':>7} {'in_t':>4} text")
    for r in res["paragraphs"]:
        text = r["text"][:20].replace("\r", " ").replace("\x07", " ")
        print(
            f"{r['i']:>3} "
            f"{r.get('page_start','?'):>4} "
            f"{r.get('y_page_start','?'):>6} "
            f"{r.get('y_text_start','?'):>6} "
            f"{r.get('y_page_ch1','?'):>6} "
            f"{str(r.get('fs','?'))[:5]:>5} "
            f"{r.get('line_spacing_rule','?'):>4} "
            f"{str(r.get('line_spacing','?'))[:7]:>7} "
            f"{('T' if r.get('in_table') else 'B'):>4} {text}"
        )
