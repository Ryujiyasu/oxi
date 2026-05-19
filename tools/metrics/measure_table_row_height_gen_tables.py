"""Measure per-row cell heights in gen_tables.docx via Word COM to derive
Word's default row height when no <w:trHeight> is specified.

Background (S110, 2026-05-19): LibreOffice beats Oxi by SSIM +0.13 on
gen_tables p.1 because Oxi compresses table rows. OOXML has no <w:trHeight>
and no <w:tblCellMar>, so Word is using default cell padding +
default lineRule=auto + default font size to derive row height. We
need to measure what that is.

Output: pipeline_data/com_measurements/gen_tables_row_heights.json

For each row of each table:
  - row_top_y, row_bottom_y (pt, page coords)
  - row_height = bottom - top
  - cell text + font + size (first cell of the row)

Pre-req: Word installed on this machine, pywin32 in current Python env.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
DOCX = REPO_ROOT / "tools" / "golden-test" / "documents" / "docx" / "gen_tables.docx"
OUT = REPO_ROOT / "pipeline_data" / "com_measurements" / "gen_tables_row_heights.json"

# Word COM constants
WD_VERT_POS_REL_TO_PAGE = 6
WD_HORIZ_POS_REL_TO_PAGE = 5


def main():
    import win32com.client

    if not DOCX.is_file():
        sys.exit(f"docx not found: {DOCX}")

    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(DOCX), ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.Repaginate()
            tables_out = []
            for ti in range(1, doc.Tables.Count + 1):
                tbl = doc.Tables(ti)
                n_rows = tbl.Rows.Count
                n_cols = tbl.Columns.Count
                rows_out = []
                for ri in range(1, n_rows + 1):
                    try:
                        row = tbl.Rows(ri)
                    except Exception as e:
                        rows_out.append({"row_idx": ri, "error": f"row access: {e}"})
                        continue
                    try:
                        cell = row.Cells(1)
                    except Exception as e:
                        rows_out.append({"row_idx": ri, "error": f"cell access: {e}"})
                        continue

                    # Get cell range vertical extents via Information(6) on
                    # start and end of cell range.
                    cell_range = cell.Range
                    start_y = doc.Range(cell_range.Start, cell_range.Start).Information(WD_VERT_POS_REL_TO_PAGE)
                    end_y = doc.Range(cell_range.End - 1, cell_range.End - 1).Information(WD_VERT_POS_REL_TO_PAGE)

                    # Row's reported Height property — None if HeightRule=auto, otherwise value in pt.
                    try:
                        height_rule = row.HeightRule
                        explicit_height_pt = row.Height
                    except Exception:
                        height_rule = None
                        explicit_height_pt = None

                    cell_text = cell.Range.Text.replace("\r", " ").replace("\x07", "").strip()[:30]
                    try:
                        font_name = cell.Range.Font.Name
                        font_size = cell.Range.Font.Size
                    except Exception:
                        font_name = None
                        font_size = None

                    rows_out.append({
                        "row_idx": ri,
                        "start_y_pt": round(start_y, 2),
                        "end_y_pt": round(end_y, 2),
                        "y_span_pt": round(end_y - start_y, 2),
                        "height_rule": height_rule,
                        "explicit_height_pt": explicit_height_pt,
                        "cell_text_first": cell_text,
                        "font_name": font_name,
                        "font_size_pt": font_size,
                    })

                # Also measure row-to-row top deltas (this is the actual
                # rendered row height in page coordinates).
                for i in range(len(rows_out) - 1):
                    a = rows_out[i]
                    b = rows_out[i + 1]
                    if "start_y_pt" in a and "start_y_pt" in b:
                        a["next_row_top_delta_pt"] = round(b["start_y_pt"] - a["start_y_pt"], 2)

                tables_out.append({
                    "table_idx": ti,
                    "n_rows": n_rows,
                    "n_cols": n_cols,
                    "rows": rows_out,
                })

            result = {
                "doc": DOCX.name,
                "n_tables": len(tables_out),
                "tables": tables_out,
            }
            OUT.parent.mkdir(parents=True, exist_ok=True)
            with OUT.open("w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            print(f"# wrote {OUT}")

            # Summary print
            for t in tables_out:
                print(f"\nTable {t['table_idx']}: {t['n_rows']} rows x {t['n_cols']} cols")
                for r in t["rows"]:
                    if "y_span_pt" in r:
                        ndt = r.get("next_row_top_delta_pt", "—")
                        print(f"  row {r['row_idx']:>2}: top={r['start_y_pt']:6.2f} bot={r['end_y_pt']:6.2f} "
                              f"span={r['y_span_pt']:5.2f}pt next_dy={ndt}  "
                              f"rule={r.get('height_rule')} explicit={r.get('explicit_height_pt')}  "
                              f"font={r.get('font_name')} {r.get('font_size_pt')}pt  "
                              f"text={r['cell_text_first']!r}")
        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
