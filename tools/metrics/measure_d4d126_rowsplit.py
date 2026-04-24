"""Focused measurement on d4d126 to confirm row-split cursor_y formula
on a 3rd real doc. Limits scan to first 10 tables only for speed."""
import json
import sys
from pathlib import Path
from statistics import median
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\ed025cbecffb_index-23.docx")
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\ed025_rowsplit.json")


def measure(doc):
    results = {"file": DOCX.name, "page_count": doc.ComputeStatistics(2), "tables": []}
    n_tables = doc.Tables.Count
    print(f"  Total tables: {n_tables}", flush=True)
    for ti in range(1, n_tables + 1):
        try:
            tbl = doc.Tables(ti)
        except Exception:
            continue
        rows = tbl.Rows.Count
        cols = tbl.Columns.Count
        # Measure ALL split rows (not just first)
        split_list = []
        for ri in range(1, rows + 1):
            try:
                row = tbl.Rows(ri)
                s = doc.Range(row.Range.Start, row.Range.Start + 1).Information(3)
                e = doc.Range(row.Range.End - 1, row.Range.End).Information(3)
                if e > s:
                    split_list.append((ri, row, s, e))
            except Exception:
                continue
        if not split_list:
            continue
        print(f"  tbl#{ti} ({rows}x{cols}) {len(split_list)} split row(s)", flush=True)
        # Process just first split row for efficiency
        ri, row, start_pg, end_pg = split_list[0]
        print(f"    split_row={ri} pg {start_pg}->{end_pg}", flush=True)

        # Measure first cell's continuation
        for ci in range(1, min(cols, 3) + 1):
            try:
                cell = row.Cells(ci)
                cell_range = cell.Range
                seen = {}
                step = max(1, (cell_range.End - cell_range.Start) // 200)
                for off in range(cell_range.Start, cell_range.End, step):
                    try:
                        r = doc.Range(off, off + 1)
                        pg = r.Information(3)
                        y = r.Information(6)
                    except Exception:
                        continue
                    if pg <= start_pg:
                        continue
                    key = (int(pg), round(y, 1))
                    if key not in seen:
                        seen[key] = off
                ys_sorted = sorted(seen.keys())
                if not ys_sorted:
                    continue
                first_pg = ys_sorted[0][0]
                page_ys = [y for p, y in ys_sorted if p == first_pg]
                filtered = [page_ys[0]]
                for y in page_ys[1:]:
                    if y - filtered[-1] < 50:
                        filtered.append(y)
                    else:
                        break
                lh = None
                if len(filtered) >= 2:
                    diffs = [filtered[i+1] - filtered[i] for i in range(len(filtered)-1)]
                    lh = round(median(diffs), 2)
                # trailing empty
                paras = cell.Range.Paragraphs
                last_p = paras(paras.Count)
                last_p_text = last_p.Range.Text.strip("\r\n\x07 \t")
                has_trailing = len(last_p_text) == 0

                # First body after table
                tbl_end = tbl.Range.End
                after_body = None
                for pi in range(1, min(doc.Paragraphs.Count + 1, 200)):
                    try:
                        p = doc.Paragraphs(pi)
                        pr = p.Range
                        if pr.Start < tbl_end:
                            continue
                        y = pr.Information(6)
                        pg = pr.Information(3)
                        if bool(pr.Information(12)):
                            continue
                        after_body = {"idx": pi, "page": int(pg), "y_pt": round(y, 2)}
                        break
                    except Exception:
                        continue

                pred = None
                resid = None
                if filtered and lh is not None and after_body:
                    last_y = filtered[-1]
                    te = 1 if has_trailing else 0
                    pred = round(last_y + lh * (1 + te), 2)
                    resid = round(after_body["y_pt"] - pred, 2)

                d = {
                    "table_idx": ti, "col": ci,
                    "continuation_lines": len(filtered),
                    "last_new_page_y": filtered[-1] if filtered else None,
                    "line_height": lh,
                    "has_trailing_empty": has_trailing,
                    "first_body_after": after_body,
                    "formula_prediction": pred,
                    "formula_residual": resid,
                }
                results["tables"].append(d)
                print(f"    cell={ci} cont={len(filtered)} lh={lh} last_y={filtered[-1] if filtered else '?'} TE={has_trailing} body={after_body['y_pt'] if after_body else '?'} pred={pred} Δ={resid}", flush=True)
            except Exception as ex:
                print(f"    cell={ci} err: {ex}", flush=True)
                continue
    return results


def main():
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(DOCX.resolve()), ReadOnly=True)
        try:
            r = measure(doc)
        finally:
            doc.Close(SaveChanges=0)
    finally:
        word.Quit()
    OUT.write_text(json.dumps(r, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
