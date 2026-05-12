"""R7.28 — Measure de6e table 5 row 12 Word rendered height.

Critical question: does Word render the vMerge=restart cell's content
in row 12 itself (so row 12 is ~205pt tall) or distribute across the
vMerge span (so row 12 is ~32pt tall)?

If row 12 is ~32pt in Word: R7.26 fix is correct; de6e cascade is upstream.
If row 12 is ~205pt in Word: vMerge=restart keeps content in restart row
  for de6e — different mechanism than 6514/a1d6/d4d126.

Strategy: for de6e t5, walk paragraphs and record per-row entry Y +
first-cell row height via Cell.RowIndex + Cell.Height.
"""
from __future__ import annotations
import os
import sys
import time
import glob
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


def find_docx(doc_id: str) -> str | None:
    candidates = glob.glob(os.path.join(REPO, "tools", "golden-test", "documents", "docx", f"{doc_id}*.docx"))
    return candidates[0] if candidates else None


def main() -> int:
    docx = find_docx("de6e32b5960b")
    if not docx:
        print("[NG] de6e not found")
        return 1
    print(f"docx: {docx}")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        time.sleep(0.4)
        n_tables = doc.Tables.Count
        print(f"top-level tables: {n_tables}")
        if n_tables < 5:
            print("[NG] table 5 missing")
            return 1
        t5 = doc.Tables(5)
        tbl_start = t5.Range.Start
        tbl_end = t5.Range.End
        print(f"table 5 range: {tbl_start}..{tbl_end}")

        # Walk paragraphs and group by row.
        rows: dict[int, dict] = {}
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            p_start = p.Range.Start
            if p_start < tbl_start or p_start >= tbl_end:
                continue
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
            start_rng = doc.Range(p_start, p_start)
            page = start_rng.Information(3)
            y = start_rng.Information(6)
            text = p.Range.Text.replace("\r", "").replace("\x07", "").strip()[:40]
            if row_num not in rows:
                rows[row_num] = {
                    "first_y": y,
                    "page": page,
                    "first_text": text,
                    "cells": [],
                    "para_count": 0,
                }
            rows[row_num]["para_count"] += 1
            rows[row_num]["last_y"] = y

        # For each row in range 8-16 (around row 12), report cell counts and
        # try to get cell height from the table's Cell(r, c).
        print(f"\nrows discovered: {sorted(rows.keys())[:20]}")
        print(f"\n{'r':>3} {'page':>4} {'first_y':>8} {'last_y':>8} {'np':>3}  first_text")
        for r in sorted(rows.keys()):
            if 8 <= r <= 16:
                info = rows[r]
                print(f"{r:>3} {info['page']:>4} {info['first_y']:>8.2f} {info.get('last_y', info['first_y']):>8.2f} {info['para_count']:>3}  {info['first_text']}")

        # Compute row-height from y-delta between consecutive rows.
        print(f"\n--- row heights (y delta) ---")
        sorted_r = sorted(rows.keys())
        for i, r in enumerate(sorted_r):
            if i + 1 < len(sorted_r):
                nxt = sorted_r[i + 1]
                if rows[r]["page"] == rows[nxt]["page"]:
                    rh = rows[nxt]["first_y"] - rows[r]["first_y"]
                    marker = " ***" if r in (11, 12, 13) else ""
                    if 8 <= r <= 16:
                        print(f"row {r}→{nxt}: {rh:+.2f}pt{marker}")

        # Now use Cell direct access for row 12 first cell height.
        print(f"\n--- direct cell access ---")
        try:
            # row 12 cell 1
            for r_idx in (10, 11, 12, 13, 14):
                try:
                    row = t5.Rows(r_idx)
                    # Word's Row.Height returns rendered height
                    rh = row.Height
                    h_rule = row.HeightRule
                    print(f"row {r_idx}: Row.Height={rh:.2f}pt rule={h_rule}")
                except Exception as e:
                    print(f"row {r_idx}: ERR {e}")
        except Exception as e:
            print(f"row access ERR: {e}")

        doc.Close(SaveChanges=False)
    finally:
        word.Quit()
    return 0


if __name__ == "__main__":
    sys.exit(main())
