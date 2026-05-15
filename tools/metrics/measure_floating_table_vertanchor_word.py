"""COM-measure Word's placement of paragraphs in vertAnchor="text" repros.

For each variant produced by `build_floating_table_vertanchor_repros.py`,
open the docx in Word COM and record Information(6) (= y in pt) for the
ANCHOR paragraph + each BODY-i trailing paragraph + each CELL-i.

This tells us:
  - Does Word push BODY-1..N past the floating table's bottom?
  - Does the answer depend on tblpY (small / mid / negative)?
  - What is the actual y of the trailing body paragraphs?

Compared to v5 (no_float baseline), the BODY-1..N positions tell us the
exact Word behavior we need to replicate.

Output: pipeline_data/ra_manual_measurements/floating_table_vertanchor_word.json
"""
from __future__ import annotations

import json
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = r"c:\Users\ryuji\oxi-main"
REPRO_DIR = r"c:\tmp"
OUT = os.path.join(
    REPO, "pipeline_data", "ra_manual_measurements",
    "floating_table_vertanchor_word.json",
)

VARIANTS = [
    "vfloat_v1_small_y",
    "vfloat_v2_small_y_tall",
    "vfloat_v3_mid_y",
    "vfloat_v4_neg_y",
    "vfloat_v5_no_float",
    "vfloat_v6_fullw_small_y",
    "vfloat_v7_fullw_tall",
    "vfloat_v8_fullw_mid_y",
    "vfloat_v9_fullw_no_horz",
    "vfloat_v10_fullw_horz_column",
    "vfloat_v11_fullw_body_is_table",
    "vfloat_v12_fullw_no_horz_body_tbl",
    "vfloat_v13_fullw_horz_missing_tblpX641",
    "vfloat_v14_fullw_horz_margin_tblpX641",
    "vfloat_v15_fullw_horz_page_tblpX2008",
    "vfloat_v16_narrow_horz_missing_tblpX641",
]


def measure(word, docx_path: str) -> dict:
    doc = word.Documents.Open(docx_path, ReadOnly=True)
    time.sleep(0.3)
    rows = []
    try:
        n_paras = doc.Paragraphs.Count
        for i in range(1, n_paras + 1):
            p = doc.Paragraphs(i)
            rng = p.Range
            # R30 fix: collapsed start range, not active end
            start_rng = doc.Range(rng.Start, rng.Start)
            page = start_rng.Information(3)  # wdActiveEndPageNumber
            y = start_rng.Information(6)     # wdVerticalPositionRelativeToPage
            x = start_rng.Information(5)     # wdHorizontalPositionRelativeToPage
            try:
                in_table = p.Range.Information(12)  # wdWithInTable
            except Exception:
                in_table = False
            text = rng.Text.replace("\r", "").replace("\n", "").strip()
            rows.append({
                "i": i,
                "page": int(page),
                "y": float(y),
                "x": float(x),
                "in_table": bool(in_table),
                "text": text[:60],
            })
    finally:
        doc.Close(SaveChanges=False)
    return {"n_paras": rows.__len__(), "paragraphs": rows}


def main() -> int:
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    summary: dict[str, dict] = {}
    try:
        for label in VARIANTS:
            docx = os.path.join(REPRO_DIR, f"{label}.docx")
            if not os.path.exists(docx):
                print(f"[skip] {label}: not found at {docx}")
                continue
            print(f"\n=== {label} ===")
            res = measure(word, docx)
            summary[label] = res
            for r in res["paragraphs"]:
                print(f"  p{r['i']:>2}  page={r['page']:>2}  y={r['y']:>7.2f}pt  "
                      f"x={r['x']:>7.2f}pt  in_table={int(r['in_table'])}  "
                      f"text={r['text']!r}")
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
