"""Focused measurement on 3a4f for 3rd-doc Path B evidence.
Scan each table quickly — only find 1-2 tables with clean splits + body-after data."""
import json
import sys
from pathlib import Path
from statistics import median
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\3a4f9fbe1a83_001620506.docx")
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\3a4f_rowsplit.json")

MAX_CELLS_PER_TABLE = 2  # only measure first 2 cells


def measure(doc, max_tables=30):
    result = {"file": DOCX.name, "tables": []}
    tbl_total = min(doc.Tables.Count, max_tables)
    print(f"Scanning first {tbl_total} of {doc.Tables.Count} tables", flush=True)
    for ti in range(1, tbl_total + 1):
        try:
            tbl = doc.Tables(ti)
        except Exception:
            continue
        rows = tbl.Rows.Count
        cols = tbl.Columns.Count
        split = None
        for ri in range(1, rows + 1):
            try:
                row = tbl.Rows(ri)
                s = doc.Range(row.Range.Start, row.Range.Start + 1).Information(3)
                e = doc.Range(row.Range.End - 1, row.Range.End).Information(3)
                if e > s:
                    split = (ri, row, s, e)
                    break
            except Exception:
                continue
        if split is None:
            continue
        ri, row, start_pg, end_pg = split
        print(f"  tbl#{ti} ({rows}x{cols}) split_row={ri} pg {start_pg}->{end_pg}", flush=True)
        # Measure only col 1 for speed
        for ci in range(1, min(cols, MAX_CELLS_PER_TABLE) + 1):
            try:
                cell = row.Cells(ci)
                cell_range = cell.Range
                # Sample sparsely
                step = max(1, (cell_range.End - cell_range.Start) // 300)
                seen = {}
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
                paras = cell.Range.Paragraphs
                last_p = paras(paras.Count)
                last_p_text = last_p.Range.Text.strip("\r\n\x07 \t")
                has_trailing = len(last_p_text) == 0

                tbl_end = tbl.Range.End
                after_body = None
                # Sparse iteration: scan up to 50 paragraphs after table_end
                for pi in range(1, min(doc.Paragraphs.Count + 1, 400)):
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

                result["tables"].append({
                    "table_idx": ti, "col": ci,
                    "continuation_lines": len(filtered),
                    "last_new_page_y": filtered[-1] if filtered else None,
                    "line_height": lh,
                    "has_trailing_empty": has_trailing,
                    "first_body_after": after_body,
                    "formula_prediction": pred,
                    "formula_residual": resid,
                })
                body_y = after_body['y_pt'] if after_body else '?'
                print(f"    cell={ci} cont={len(filtered)} lh={lh} last_y={filtered[-1] if filtered else '?'} TE={has_trailing} body={body_y} pred={pred} Δ={resid}", flush=True)
                # stop after first col with body data
                if after_body is not None:
                    break
            except Exception as ex:
                print(f"    cell={ci} err: {ex}", flush=True)
    return result


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
