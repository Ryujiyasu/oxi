"""COM measurement for table_row_height_drift minimal repro variants.

For each .docx in this directory, opens Word and reads:
- Y of the marker paragraph above the table (Y0)
- Y of the (auto-generated) paragraph after the table (Y_after)
- Y of each character in the cell's single paragraph via Range.Characters

The cell content is N copies of "あ" separated by <w:br/> soft line breaks.
Word lays them out one per line. We pick the FIRST character per line by
detecting Y jumps in the per-character Y stream.

Computes:
- table_height_pt = Y_after - Y0
- per_line_y[]    = de-duplicated character Ys
- per_line_delta  = consecutive diffs (the snapped advance)

Run on Windows with: pip install pywin32
"""
from __future__ import annotations

import json
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


def measure_one(word, docx_path: Path) -> dict:
    doc = word.Documents.Open(str(docx_path), ReadOnly=True)
    time.sleep(0.4)
    try:
        result: dict = {"file": docx_path.name}

        # Find before/after Y via paragraph scan
        n_paras = doc.Paragraphs.Count
        before_y = None
        after_y = None
        in_table_run = False
        for pi in range(1, n_paras + 1):
            p = doc.Paragraphs(pi)
            in_tbl = bool(p.Range.Information(WD_WITHIN_TABLE))
            y = p.Range.Information(WD_Y_PAGE)
            if not in_tbl and not in_table_run and before_y is None and pi < n_paras:
                before_y = y
            elif in_tbl:
                in_table_run = True
            elif in_table_run and not in_tbl and after_y is None:
                after_y = y
                break
        result["before_para_y_pt"] = before_y
        result["after_para_y_pt"] = after_y
        if before_y is not None and after_y is not None:
            result["table_height_pt"] = round(after_y - before_y, 4)
        else:
            result["table_height_pt"] = None

        if doc.Tables.Count == 0:
            result["error"] = "no table"
            return result
        t = doc.Tables(1)
        cell = t.Cell(1, 1)

        # Walk every character in the cell's range and record their Y.
        # First-of-line characters are detected by Y change.
        chars = cell.Range.Characters
        n_chars = chars.Count
        line_ys: list[float] = []
        last_y: float | None = None
        for ci in range(1, n_chars + 1):
            c = chars(ci)
            try:
                y = float(c.Information(WD_Y_PAGE))
            except Exception:
                continue
            # Skip the trailing cell-end-marker (\x07), which sometimes
            # reports a junk Y. We accept any sane Y > 0.
            if y <= 0:
                continue
            if last_y is None or abs(y - last_y) > 0.01:
                # Y changed → new line (or first char)
                if last_y is None or y > last_y:
                    line_ys.append(y)
                    last_y = y

        result["line_count"] = len(line_ys)
        result["line_ys_pt"] = line_ys
        # Per-line delta: consecutive diffs
        deltas = [round(line_ys[i + 1] - line_ys[i], 4) for i in range(len(line_ys) - 1)]
        result["line_delta_pt"] = deltas
        # Sum of deltas
        result["sum_deltas_pt"] = round(sum(deltas), 4)
        # First and last line Y
        if line_ys:
            result["first_line_y_pt"] = line_ys[0]
            result["last_line_y_pt"] = line_ys[-1]
            result["span_pt"] = round(line_ys[-1] - line_ys[0], 4)
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
    return 0


if __name__ == "__main__":
    sys.exit(main())
