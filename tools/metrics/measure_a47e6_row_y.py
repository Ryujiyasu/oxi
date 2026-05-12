"""Day 33 part 68 R7.22 session 4 — Measure Word's actual per-row Y positions
for a47e6 table 1 (the form table).

Compares Word COM Range.Information(6) for the START of each row vs Oxi
TBL_DUMP entry_cursor_y. Per-row diff pinpoints where the cumulative
over-pump (Oxi pushing 1.4pt past pgBot for pi=2) really comes from.

For each row, take the first character of the first cell and read its Y.

Run: python tools/metrics/measure_a47e6_row_y.py
Output: pipeline_data/a47e6_row_y_compare.csv
"""

from __future__ import annotations
import sys
import os
import re
import csv
import time
import subprocess
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                    "a47e6c6b2ca1_order_08.docx")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")


def measure_word_rows() -> list[dict]:
    """Use Word COM to walk every paragraph and identify table 1 rows.

    Information(11) = wdStartOfRangeRowNumber (1-based row index within table)
    Information(12) = wdStartOfRangeColumnNumber
    Information(13) = wdActiveEndAdjustedPageNumber? — not what we want
    Use Range.Tables to detect if paragraph is in a table.
    """
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    rows: dict[int, dict] = {}
    try:
        doc = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
        time.sleep(0.3)

        n_paras = doc.Paragraphs.Count
        print(f"Total paragraphs: {n_paras}")

        # First, find tables in the doc and use table 1 specifically.
        # Iterate paragraphs and check if they belong to table 1.
        tbl1 = doc.Tables(1)
        tbl1_start = tbl1.Range.Start
        tbl1_end = tbl1.Range.End

        # For each paragraph in table 1 range, detect ROW transitions by
        # tracking row_num via Cell.RowIndex (works even when nested or
        # vMerged). Information(11) doesn't reliably reflect outer row.
        prev_row_signature = None
        for i in range(1, n_paras + 1):
            p = doc.Paragraphs(i)
            p_start = p.Range.Start
            if p_start < tbl1_start or p_start >= tbl1_end:
                continue
            # Skip nested table paragraphs by checking tables nesting level
            try:
                nesting = p.Range.Tables.Count
            except Exception:
                nesting = 0
            # nesting == 1 means in outer table 1, nesting >= 2 in nested
            if nesting != 1:
                continue
            # Get outer row via Cell.RowIndex
            try:
                cell = p.Range.Cells(1)
                row_num = cell.RowIndex
            except Exception:
                continue
            if row_num in rows:
                continue
            start_rng = doc.Range(p_start, p_start)
            page = start_rng.Information(3)
            y = start_rng.Information(6)
            text = p.Range.Text.replace("\r", "").replace("\x07", "")[:30]
            rows[row_num] = {
                "row_idx": row_num - 1,
                "word_page": page,
                "word_y": round(y, 3),
                "first_text": text,
            }

        doc.Close(SaveChanges=False)
    finally:
        word.Quit()
    return sorted(rows.values(), key=lambda r: r["row_idx"])


def measure_oxi_rows() -> list[dict]:
    """Run OXI_DUMP_TABLE and parse table 1's entry_cursor_y per row."""
    env = os.environ.copy()
    env["OXI_DUMP_TABLE"] = "1"
    proc = subprocess.run(
        [GDI, os.path.abspath(DOCX), os.path.join(REPO, "nul")],
        env=env,
        capture_output=True,
        text=False,
    )
    text = proc.stderr.decode("utf-8", errors="replace")

    for f in os.listdir(REPO):
        if f.startswith("nul_p") and f.endswith(".png"):
            try:
                os.remove(os.path.join(REPO, f))
            except OSError:
                pass

    # Parse only TOP-LEVEL row entries (= table 1 rows). Strategy: track
    # whether we're in nested table by counting "row=0 entry" occurrences:
    # the FIRST row=0 entry is table 1, subsequent row=0 entries before a
    # "row=N pre_correction row_height=" higher than row 0's are nested.
    # Simpler: top-level table 1 rows have entries with cy > 50 (page-y),
    # nested rows have cy < 100 typically. Use cy > 50 + first occurrence
    # of each row_idx in top-level as filter.
    rows = []
    seen_outer_rows = set()
    for line in text.splitlines():
        m = re.search(
            r"\[TBL_DUMP\] row=(\d+) entry_cursor_y=([\d.]+) row_height_pre=([\d.]+).*?trHeight=([\d.]+) rule=(\S+) n_cells=(\d+)",
            line,
        )
        if m:
            r_idx = int(m.group(1))
            cy = float(m.group(2))
            rh = float(m.group(3))
            trh = float(m.group(4))
            rule = m.group(5)
            n_cells = int(m.group(6))
            # Outer table 1 rows: appear first; nested ones use cy < 100 etc.
            # Use cy > 50 to filter (nested rows have small cy because relative).
            if cy > 50 and r_idx not in seen_outer_rows:
                seen_outer_rows.add(r_idx)
                rows.append({
                    "row_idx": r_idx,
                    "oxi_cy": cy,
                    "oxi_rh_pre": rh,
                    "trh": trh,
                    "rule": rule,
                    "n_cells": n_cells,
                })
    return rows


def main() -> int:
    print("Measuring Word per-row Y...")
    word_rows = measure_word_rows()
    print(f"\nMeasuring Oxi per-row entry_cursor_y...")
    oxi_rows = measure_oxi_rows()
    print(f"Word: {len(word_rows)} rows, Oxi: {len(oxi_rows)} rows\n")

    # Cross-compare by row_idx
    results = []
    print(f"{'row':>3} {'word_pg':>7} {'word_y':>8} {'oxi_cy':>8} {'delta':>8} {'oxi_rh':>7} {'trh':>6} {'rule':>10} text")
    print("-" * 100)
    for w in word_rows:
        match = next((o for o in oxi_rows if o["row_idx"] == w["row_idx"]), None)
        if match:
            delta = match["oxi_cy"] - w["word_y"]
            print(f"{w['row_idx']:>3} {w['word_page']:>7} {w['word_y']:>8.2f} "
                  f"{match['oxi_cy']:>8.2f} {delta:>+7.2f} {match['oxi_rh_pre']:>7.2f} "
                  f"{match['trh']:>6.2f} {match['rule']:>10} {w['first_text']!r}")
            results.append({**w, **match, "delta_y": round(delta, 3)})
        else:
            print(f"{w['row_idx']:>3} {w['word_page']:>7} {w['word_y']:>8.2f} {'(no oxi)':>17} {w['first_text']!r}")

    # Compute per-row increment (row_h = next.y - this.y for Word)
    print(f"\nPer-row INCREMENT comparison (row_h = next row entry - this row entry):")
    print(f"{'row':>3} {'word_rh':>8} {'oxi_rh':>8} {'delta_rh':>8}")
    for i in range(len(results) - 1):
        w_rh = results[i+1]["word_y"] - results[i]["word_y"]
        o_rh = results[i+1]["oxi_cy"] - results[i]["oxi_cy"]
        d_rh = o_rh - w_rh
        print(f"{i:>3} {w_rh:>8.2f} {o_rh:>8.2f} {d_rh:>+7.2f}")

    # CSV
    out_csv = os.path.join(REPO, "pipeline_data", "a47e6_row_y_compare.csv")
    os.makedirs(os.path.dirname(out_csv), exist_ok=True)
    if results:
        with open(out_csv, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(results[0].keys()))
            w.writeheader()
            w.writerows(results)
        print(f"\nCSV: {out_csv}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
