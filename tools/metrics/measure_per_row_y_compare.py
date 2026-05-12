"""Day 33 part 69 R7.23 — Generalized per-row Y comparison (Word vs Oxi)
for any doc with a top-level table.

Generalizes R7.22 session 4's a47e6-specific tool. For each top-level
table in the document, captures Word's actual row Y positions via
Cell.RowIndex (handles vMerge/gridSpan) and Oxi's TBL_DUMP
entry_cursor_y. Reports per-row delta to identify systematic
over-pump patterns across the cluster.

Usage:
  python tools/metrics/measure_per_row_y_compare.py <doc_id>
  python tools/metrics/measure_per_row_y_compare.py 6514f214e482

Output: pipeline_data/per_row_y_<doc_id>.csv
"""

from __future__ import annotations
import sys
import os
import re
import csv
import time
import glob
import subprocess
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")


def find_docx(doc_id: str) -> str | None:
    candidates = glob.glob(os.path.join(REPO, "tools", "golden-test", "documents", "docx", f"{doc_id}*.docx"))
    return candidates[0] if candidates else None


def measure_word(docx: str) -> list[dict]:
    """For each top-level table, capture each row's start Y via the
    first paragraph that lives in that row (Cell.RowIndex)."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    out = []
    try:
        doc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        time.sleep(0.3)

        # Iterate all top-level tables
        n_tables = doc.Tables.Count
        print(f"  Word: {n_tables} top-level tables")

        # For each top-level table, get its range bounds, then find row
        # entry Y for each row by walking paragraphs.
        for t_idx in range(1, n_tables + 1):
            tbl = doc.Tables(t_idx)
            # Skip nested tables (have parent)
            try:
                parent_tbl = tbl.Parent.InRange(doc.Range(0, doc.Range.End - 1))
                # Heuristic: top-level table's range start should be at doc-level
                if hasattr(tbl, "Range") and tbl.Range.Tables.Count > 1:
                    # this is nested
                    continue
            except Exception:
                pass

            tbl_start = tbl.Range.Start
            tbl_end = tbl.Range.End

            # Walk paragraphs once, mark first per (table, row)
            seen_rows: set[int] = set()
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(i)
                p_start = p.Range.Start
                if p_start < tbl_start or p_start >= tbl_end:
                    continue
                # Outer table only (this table)
                try:
                    if p.Range.Tables.Count != 1:
                        continue
                    if p.Range.Tables(1).Range.Start != tbl_start:
                        continue
                except Exception:
                    continue
                try:
                    cell = p.Range.Cells(1)
                    row_num = cell.RowIndex
                except Exception:
                    continue
                if row_num in seen_rows:
                    continue
                seen_rows.add(row_num)
                start_rng = doc.Range(p_start, p_start)
                page = start_rng.Information(3)
                y = start_rng.Information(6)
                text = p.Range.Text.replace("\r", "").replace("\x07", "")[:30]
                out.append({
                    "table_idx": t_idx,
                    "row_idx": row_num - 1,
                    "word_page": page,
                    "word_y": round(y, 3),
                    "first_text": text,
                })
            print(f"    table {t_idx}: {len(seen_rows)} rows")

        doc.Close(SaveChanges=False)
    finally:
        word.Quit()
    return out


def measure_oxi(docx: str) -> list[dict]:
    """Parse OXI_DUMP_TABLE output for top-level row entry_cursor_y values."""
    env = os.environ.copy()
    env["OXI_DUMP_TABLE"] = "1"
    proc = subprocess.run(
        [GDI, os.path.abspath(docx), os.path.join(REPO, "nul")],
        env=env, capture_output=True, text=False,
    )
    text = proc.stderr.decode("utf-8", errors="replace")

    for f in os.listdir(REPO):
        if f.startswith("nul_p") and f.endswith(".png"):
            try:
                os.remove(os.path.join(REPO, f))
            except OSError:
                pass

    # Parse entries; top-level rows have larger cy values (>50pt).
    # Nested table rows reset cy to small values (<60pt typical).
    # Track top-level rows by detecting "row=0 entry" patterns; each
    # top-level table starts a new row=0. Use cy threshold: cy > 50 AND
    # this is the first row=0 in a sequence.
    entries = []
    for line in text.splitlines():
        m = re.search(
            r"\[TBL_DUMP\] row=(\d+) entry_cursor_y=([\d.]+) row_height_pre=([\d.]+).*?trHeight=([\d.]+) rule=(\S+)",
            line,
        )
        if m:
            entries.append({
                "row_idx": int(m.group(1)),
                "cy": float(m.group(2)),
                "rh_pre": float(m.group(3)),
                "trh": float(m.group(4)),
                "rule": m.group(5),
            })

    # Identify top-level table boundaries: a top-level row=0 has cy > 50
    # AND the next row=N (for N>=1) sequence belongs to that table. When
    # we see another row=0 with cy<60, it's a nested table inside.
    top_level_tables = []  # list of list of row entries
    current_table = []
    current_table_id = 0

    for entry in entries:
        # Heuristic: cy > 30 and (current is empty OR row_idx == 0 starting fresh)
        # Top-level tables can start near page top (e.g., 31420af pg_top ~= 36)
        if entry["row_idx"] == 0 and entry["cy"] > 30:
            # New top-level table starts
            if current_table:
                top_level_tables.append(current_table)
            current_table = []
            current_table_id += 1
            entry["table_idx"] = current_table_id
            current_table.append(entry)
        elif entry["cy"] < 30 and entry["row_idx"] == 0:
            # Nested table starts; skip
            continue
        else:
            # Continuation row of current table (only if not nested)
            if current_table:
                # Check if this is contiguous with last entry of current_table
                if entry["row_idx"] > current_table[-1]["row_idx"]:
                    entry["table_idx"] = current_table_id
                    current_table.append(entry)
                # else: nested or different
    if current_table:
        top_level_tables.append(current_table)

    out = []
    for tbl in top_level_tables:
        for e in tbl:
            out.append(e)
    return out


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: measure_per_row_y_compare.py <doc_id>")
        return 1
    doc_id = sys.argv[1]
    docx = find_docx(doc_id)
    if not docx:
        print(f"[NG] doc not found: {doc_id}")
        return 1
    print(f"\n=== {doc_id} ===")
    print(f"  docx: {docx}")
    print(f"\n  measuring Word...")
    word_rows = measure_word(docx)
    print(f"  measuring Oxi...")
    oxi_rows = measure_oxi(docx)
    print(f"\nWord: {len(word_rows)} rows, Oxi: {len(oxi_rows)} rows\n")

    # Group by table
    from collections import defaultdict
    word_by_tbl = defaultdict(list)
    for r in word_rows:
        word_by_tbl[r["table_idx"]].append(r)
    oxi_by_tbl = defaultdict(list)
    for r in oxi_rows:
        oxi_by_tbl[r["table_idx"]].append(r)

    results = []
    for t_idx in sorted(set(word_by_tbl.keys()) | set(oxi_by_tbl.keys())):
        w_rows = word_by_tbl.get(t_idx, [])
        o_rows = oxi_by_tbl.get(t_idx, [])
        print(f"=== Table {t_idx}: Word {len(w_rows)} rows, Oxi {len(o_rows)} rows ===")
        # Per-row increment (= row_h_rendered)
        if len(w_rows) > 1 and len(o_rows) > 1:
            print(f"  {'r':>2} {'w_y':>8} {'o_cy':>8} {'r_w':>7} {'r_o':>7} {'d_r':>7}")
            for i in range(min(len(w_rows), len(o_rows))):
                w = w_rows[i]
                o = o_rows[i]
                # row height = next row entry - this row entry
                if i + 1 < min(len(w_rows), len(o_rows)):
                    w_rh = w_rows[i+1]["word_y"] - w["word_y"]
                    o_rh = o_rows[i+1]["cy"] - o["cy"]
                    # Skip transitions across page boundaries (negative or
                    # very-large positive deltas). Both must be positive
                    # AND on same page (heuristic: |rh| < 800pt).
                    if w_rh < 0 or o_rh < 0 or abs(w_rh) > 800 or abs(o_rh) > 800:
                        # Page-boundary crossing — skip from analysis
                        continue
                    d_rh = o_rh - w_rh
                    marker = " ★" if abs(d_rh) > 2 else ""
                    print(f"  {i:>2} {w['word_y']:>8.2f} {o['cy']:>8.2f} {w_rh:>7.2f} {o_rh:>7.2f} {d_rh:>+7.2f}{marker}")
                    results.append({
                        "table": t_idx, "row": i,
                        "word_y": w["word_y"], "oxi_cy": o["cy"],
                        "word_rh": w_rh, "oxi_rh": o_rh, "delta_rh": round(d_rh, 3),
                    })
                else:
                    print(f"  {i:>2} {w['word_y']:>8.2f} {o['cy']:>8.2f} (last)")
        elif w_rows or o_rows:
            print(f"  (too few rows for increment comparison)")

    out_csv = os.path.join(REPO, "pipeline_data", f"per_row_y_{doc_id}.csv")
    os.makedirs(os.path.dirname(out_csv), exist_ok=True)
    if results:
        with open(out_csv, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(results[0].keys()))
            w.writeheader()
            w.writerows(results)
        print(f"\nCSV: {out_csv}")
    # Summary
    total_pump = sum(r["delta_rh"] for r in results)
    print(f"\nTotal cumulative Oxi over-pump: {total_pump:+.2f}pt across {len(results)} measured rows")
    return 0


if __name__ == "__main__":
    sys.exit(main())
