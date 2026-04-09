"""COM measurement for table_border_overhead minimal repro variants.

For each .docx in this directory, opens Word, reads:
- Y of the marker paragraph above the table   (= table top reference)
- Y of the marker paragraph after the table   (auto-inserted by Word at EOF)
- Y of each row's first cell content          (cell.Range.Information(6))
- Row.Height + Row.HeightRule                  (declared, not actual)
- Border LineWidth for top/bot/left/right/insideH

Computes:
- table_height_pt = Y(after_marker) - Y(before_marker)
- per_row_pt      = Y(row_i+1) - Y(row_i)         (last row uses after_marker)

Dumps everything to measurements.json.

Run on Windows with: pip install pywin32
"""
from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 not installed. Run: pip install pywin32", file=sys.stderr)
    sys.exit(1)

sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]

HERE = Path(__file__).resolve().parent
WD_Y_PAGE = 6  # wdVerticalPositionRelativeToPage
WD_WITHIN_TABLE = 12  # wdWithInTable

BORDER_INDICES = [
    (-1, "top"),
    (-2, "left"),
    (-3, "bottom"),
    (-4, "right"),
    (-5, "insideH"),
    (-6, "insideV"),
]


def measure_one(word, docx_path: Path) -> dict:
    doc = word.Documents.Open(str(docx_path), ReadOnly=True)
    time.sleep(0.4)
    try:
        result: dict = {"file": docx_path.name}

        # Find paragraphs flanking the (only) table.
        # generate.py inserts a marker paragraph BEFORE the table; Word always
        # has an implicit empty paragraph AFTER the table.
        n_paras = doc.Paragraphs.Count
        before_y = None
        after_y = None
        in_table_run = False
        for pi in range(1, n_paras + 1):
            p = doc.Paragraphs(pi)
            in_tbl = bool(p.Range.Information(WD_WITHIN_TABLE))
            y = p.Range.Information(WD_Y_PAGE)
            if not in_tbl and not in_table_run and before_y is None and pi < n_paras:
                # tentative "before" — keep updating until we hit a table-paragraph
                before_y = y
                before_idx = pi
            elif in_tbl:
                in_table_run = True
            elif in_table_run and not in_tbl and after_y is None:
                after_y = y
                after_idx = pi
                break
        result["before_para_y_pt"] = before_y
        result["after_para_y_pt"] = after_y
        if before_y is not None and after_y is not None:
            result["table_height_pt"] = round(after_y - before_y, 4)
        else:
            result["table_height_pt"] = None

        # Tables
        if doc.Tables.Count == 0:
            result["error"] = "no table"
            return result
        t = doc.Tables(1)
        result["row_count"] = t.Rows.Count

        # Per-row top Y (use first cell's Range.Information(6))
        rows: list[dict] = []
        for ri in range(1, t.Rows.Count + 1):
            cell = t.Cell(ri, 1)
            top_y = cell.Range.Information(WD_Y_PAGE)
            try:
                declared_h = float(t.Rows(ri).Height)
            except Exception:
                declared_h = None
            try:
                hr = int(t.Rows(ri).HeightRule)
            except Exception:
                hr = None
            rows.append({
                "row": ri,
                "cell_top_y_pt": top_y,
                "declared_height_pt": declared_h,
                "height_rule": hr,
            })
        # Per-row height = next row top − this row top; last row uses after_y
        for i, r in enumerate(rows):
            if i + 1 < len(rows):
                r["row_height_pt"] = round(rows[i + 1]["cell_top_y_pt"] - r["cell_top_y_pt"], 4)
            elif after_y is not None:
                r["row_height_pt"] = round(after_y - r["cell_top_y_pt"], 4)
            else:
                r["row_height_pt"] = None
        result["rows"] = rows

        # Borders
        borders: dict = {}
        for idx, name in BORDER_INDICES:
            try:
                b = t.Borders(idx)
                borders[name] = {
                    "line_style": int(b.LineStyle),
                    "line_width_pt": float(b.LineWidth),
                }
            except Exception as e:
                borders[name] = {"error": repr(e)}
        result["borders"] = borders

        return result
    finally:
        doc.Close(SaveChanges=False)


def main() -> int:
    docx_files = sorted(HERE.glob("*.docx"))
    if not docx_files:
        print(f"No .docx in {HERE}. Run generate.py first.", file=sys.stderr)
        return 1

    print(f"Measuring {len(docx_files)} variants via Word COM...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        results = []
        for f in docx_files:
            print(f"  {f.name}")
            try:
                results.append(measure_one(word, f))
            except Exception as e:
                results.append({"file": f.name, "error": repr(e)})
    finally:
        word.Quit()

    out = HERE / "measurements.json"
    out.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nWrote {out}")
    print("Next: inspect measurements.json and derive border_overhead formula in analysis.md")
    return 0


if __name__ == "__main__":
    sys.exit(main())
