"""Measure Word's per-row dimensions for 459f05 floating tables.

Uses COM:
- doc.Tables(i).Rows(j).Cells(1).Range.Information(3)  -> page index
- doc.Range(start_pos, start_pos).Information(6)       -> rendered top-Y of row's first cell
- Row.Height + Row.HeightRule                          -> declared height (may not match rendered)

For each row we capture (table_idx, row_idx, page, top_y, height_pt, cells, first_text).
By taking adjacent row top_y values we infer rendered row height (handles atLeast +
auto rows that Word stretched).

Output: pipeline_data/459f05_word_table_rows.json
"""
import os, sys, time, json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = os.path.abspath(
    r"tools\golden-test\documents\docx\459f05f1e877_kyodokenkyuyoushiki01.docx"
)

PT_PER_TWIP = 1.0 / 20.0


def collapsed_start_info(doc, rng, info_id):
    """R30 fix: query Information() at a zero-length range starting at rng.Start
    to get the row's true top-of-page Y, not the active-end position."""
    s = rng.Start
    return doc.Range(s, s).Information(info_id)


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(DOCX, ReadOnly=True)
        time.sleep(0.6)

        rows_data = []
        print(f"Doc has {doc.Tables.Count} tables")
        for ti in range(1, doc.Tables.Count + 1):
            t = doc.Tables(ti)
            print(f"Table {ti}: iterating via t.Range.Cells (RowIndex grouping)")
            # Iterate cells via Range.Cells (works with vMerge) and group by RowIndex.
            # Track first cell encountered for each row to get top_y.
            seen_rows = {}  # row_idx -> first cell info
            try:
                cells = list(t.Range.Cells)
            except Exception:
                # Some Word builds expose Cells as a collection not list-iterable.
                cells = []
                for ci in range(1, t.Range.Cells.Count + 1):
                    cells.append(t.Range.Cells(ci))
            for c in cells:
                try:
                    ri = int(c.RowIndex)
                    rng = c.Range
                    page = int(collapsed_start_info(doc, rng, 3))
                    top_y = float(collapsed_start_info(doc, rng, 6))
                    if ri not in seen_rows:
                        # Try to get declared height from the cell's row
                        # (may still raise on vMerge tables; tolerate)
                        decl_h = -1.0
                        h_rule = -1
                        try:
                            row_obj = c.Row
                            decl_h = float(row_obj.Height)
                            h_rule = int(row_obj.HeightRule)
                        except Exception:
                            pass
                        first_text = rng.Text.replace("\r", " ").replace("\x07", "")[:32]
                        seen_rows[ri] = {
                            "table_idx": ti,
                            "row_idx": ri,
                            "page": page,
                            "top_y_pt": round(top_y, 3),
                            "declared_height_pt": round(decl_h, 3),
                            "height_rule": h_rule,
                            "n_cells": 1,
                            "first_text": first_text,
                        }
                    else:
                        seen_rows[ri]["n_cells"] += 1
                except Exception as e:
                    pass
            # Append in row order
            for ri in sorted(seen_rows.keys()):
                rows_data.append(seen_rows[ri])

        # Compute per-row rendered height: difference of consecutive top_y values
        # (or to next-row start; on page boundary, height = page_bottom - top_y)
        for ti in range(1, doc.Tables.Count + 1):
            same = [r for r in rows_data if r.get("table_idx") == ti and "error" not in r]
            for idx in range(len(same)):
                cur = same[idx]
                if idx + 1 < len(same):
                    nxt = same[idx + 1]
                    if nxt["page"] == cur["page"]:
                        cur["rendered_height_pt"] = round(nxt["top_y_pt"] - cur["top_y_pt"], 3)
                    else:
                        cur["rendered_height_pt"] = None  # page break between rows
                        cur["next_row_on_page"] = nxt["page"]
                        cur["next_row_top_y"] = nxt["top_y_pt"]
                else:
                    cur["rendered_height_pt"] = None  # last row of table

        # Print summary
        print("\nPer-row summary:")
        print(f"{'tbl':>3} {'row':>3} {'pg':>3} {'top_y':>8} {'rendH':>8} {'declH':>8} {'rule':>4} {'cells':>5} {'text':>32}")
        print("-" * 92)
        for r in rows_data:
            if "error" in r:
                continue
            rh = r.get("rendered_height_pt")
            rh_str = f"{rh:.2f}" if isinstance(rh, (int, float)) else "(page-break)"
            decl_h = r.get("declared_height_pt", -1)
            decl_str = f"{decl_h:.2f}" if decl_h >= 0 else "auto"
            print(f"{r['table_idx']:>3} {r['row_idx']:>3} {r['page']:>3} "
                  f"{r['top_y_pt']:>8.2f} {rh_str:>8} {decl_str:>8} "
                  f"{r.get('height_rule', -1):>4} {r['n_cells']:>5} "
                  f"{r['first_text']!r:>32}")

        out = "pipeline_data/459f05_word_table_rows.json"
        os.makedirs(os.path.dirname(out), exist_ok=True)
        with open(out, "w", encoding="utf-8") as f:
            json.dump(rows_data, f, ensure_ascii=False, indent=2)
        print(f"\nSaved: {out}")

        doc.Close(False)
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
