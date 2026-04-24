"""Extended row-split measurement v3: add b5f706 / 1636d28e to target set,
also measure effective line_height (y2 - y1 of adjacent continuation lines)
to validate formula: body_y = last_new_page_y + lh × (1 + has_trailing_empty).

For each splitting table, extract:
- split_row, row_start_page, row_end_page
- dominant cell continuation lines (via y-sampling)
- first/last new-page y
- effective line_height (median of y-diffs between adjacent continuation lines)
- first body paragraph y after table
- has_trailing_empty (probe: cell's last paragraph has no runs)
"""
import json
import sys
from pathlib import Path
from statistics import median
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx")
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\rowsplit_boundaries_v3.json")

TARGETS = [
    "d77a58485f16_20240705_resources_data_outline_08.docx",
    "e3c545fac7a7_LOD_Handbook.docx",
    "ed025cbecffb_index-23.docx",
]


def info_start_page(doc, rng):
    try:
        return doc.Range(rng.Start, rng.Start + 1).Information(3)
    except Exception:
        return None


def info_end_page(doc, rng):
    try:
        return doc.Range(rng.End - 1, rng.End).Information(3)
    except Exception:
        return None


def measure_cell_continuation(doc, cell, row_start_page):
    cell_range = cell.Range
    seen_ys = {}
    for off in range(cell_range.Start, cell_range.End):
        try:
            r = doc.Range(off, off + 1)
            pg = r.Information(3)
            y = r.Information(6)
        except Exception:
            continue
        if pg <= row_start_page:
            continue
        key = (int(pg), round(y, 1))
        if key not in seen_ys:
            seen_ys[key] = off
    if not seen_ys:
        return 0, None, None, None, None
    ys_sorted = sorted(seen_ys.keys())
    first_pg = ys_sorted[0][0]
    page_ys = [y for p, y in ys_sorted if p == first_pg]
    filtered = [page_ys[0]]
    for y in page_ys[1:]:
        if y - filtered[-1] < 50:
            filtered.append(y)
        else:
            break
    # Derive effective line_height from filtered sequence
    lh = None
    if len(filtered) >= 2:
        diffs = [filtered[i+1] - filtered[i] for i in range(len(filtered)-1)]
        lh = round(median(diffs), 2)
    # Also find LAST page if multi-page continuation
    last_pg = ys_sorted[-1][0]
    last_pg_ys = sorted(set(y for p, y in ys_sorted if p == last_pg))
    last_pg_last_y = last_pg_ys[-1] if last_pg_ys else None
    # Filter last_pg_ys similarly
    last_pg_filtered = [last_pg_ys[0]] if last_pg_ys else []
    for y in last_pg_ys[1:] if last_pg_ys else []:
        if y - last_pg_filtered[-1] < 50:
            last_pg_filtered.append(y)
        else:
            break
    last_pg_clean_last = last_pg_filtered[-1] if last_pg_filtered else None
    return len(filtered), filtered[0], filtered[-1], lh, (last_pg, last_pg_clean_last)


def measure_doc(word, docx_path: Path):
    result = {"file": docx_path.name, "tables": []}
    doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
    try:
        result["page_count"] = doc.ComputeStatistics(2)
        n_tables = doc.Tables.Count
        for ti in range(1, n_tables + 1):
            try:
                tbl = doc.Tables(ti)
            except Exception:
                continue

            rows = tbl.Rows.Count
            cols = tbl.Columns.Count

            split_row_idx = None
            split_row = None
            row_start_page = None
            row_end_page = None
            for ri in range(1, rows + 1):
                try:
                    row = tbl.Rows(ri)
                    r_start = info_start_page(doc, row.Range)
                    r_end = info_end_page(doc, row.Range)
                    if r_start is not None and r_end is not None and r_end > r_start:
                        split_row_idx = ri
                        split_row = row
                        row_start_page = r_start
                        row_end_page = r_end
                        break
                except Exception:
                    continue

            if split_row_idx is None:
                continue

            cell_data = []
            trailing_probe = []
            for ci in range(1, cols + 1):
                try:
                    cell = split_row.Cells(ci)
                    cont_count, first_y, last_y, lh, last_pg_info = \
                        measure_cell_continuation(doc, cell, row_start_page)
                    if cont_count > 0:
                        # Probe trailing-empty: last paragraph of cell
                        paras = cell.Range.Paragraphs
                        last_p = paras(paras.Count)
                        last_p_text = last_p.Range.Text.strip("\r\n\x07 \t")
                        has_trailing = len(last_p_text) == 0
                        last_pg, last_pg_y = last_pg_info
                        cell_data.append({
                            "col": ci,
                            "continuation_line_count": cont_count,
                            "first_new_page_y": first_y,
                            "last_new_page_y": last_y,
                            "last_page": int(last_pg) if last_pg else None,
                            "last_page_last_y": last_pg_y,
                            "line_height": lh,
                            "paragraphs": cell.Range.Paragraphs.Count,
                            "has_trailing_empty": has_trailing,
                        })
                except Exception as e:
                    print(f"    cell {ci} err: {e}")
                    continue

            if not cell_data:
                continue

            tbl_end = tbl.Range.End
            after_body = None
            for pi in range(1, doc.Paragraphs.Count + 1):
                try:
                    p = doc.Paragraphs(pi)
                    pr = p.Range
                    if pr.Start < tbl_end:
                        continue
                    y = pr.Information(6)
                    pg = pr.Information(3)
                    in_table = bool(pr.Information(12))
                    if in_table:
                        continue
                    after_body = {
                        "idx": pi,
                        "page": int(pg),
                        "y_pt": round(y, 2),
                    }
                    break
                except Exception:
                    continue

            dominant_cell = max(cell_data, key=lambda c: c["continuation_line_count"])

            # Compute formula prediction for validation
            pred = None
            resid = None
            if dominant_cell["last_page_last_y"] is not None and dominant_cell["line_height"] is not None and after_body:
                last_y = dominant_cell["last_page_last_y"]
                lh = dominant_cell["line_height"]
                te = 1 if dominant_cell["has_trailing_empty"] else 0
                pred = round(last_y + lh * (1 + te), 2)
                actual = after_body["y_pt"]
                resid = round(actual - pred, 2)

            t_info = {
                "table_idx": ti,
                "rows": rows,
                "cols": cols,
                "split_row": split_row_idx,
                "row_start_page": row_start_page,
                "row_end_page": row_end_page,
                "multi_page": (row_end_page - row_start_page) > 1,
                "cells_with_continuation": cell_data,
                "dominant_cell": dominant_cell,
                "first_body_after": after_body,
                "formula_prediction": pred,
                "formula_residual": resid,
            }
            result["tables"].append(t_info)
            mp = "MULTI" if t_info["multi_page"] else "1pg"
            te = "TE=1" if dominant_cell["has_trailing_empty"] else "TE=0"
            print(f"  Table {ti} ({rows}x{cols}) split={split_row_idx} {mp} {te} "
                  f"cont={dominant_cell['continuation_line_count']} "
                  f"lh={dominant_cell['line_height']} "
                  f"last_y={dominant_cell['last_page_last_y']} body={after_body['y_pt'] if after_body else '?'} "
                  f"pred={pred} Δ={resid}")
    finally:
        doc.Close(SaveChanges=0)
    return result


def main():
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    all_results = {}
    try:
        for doc_name in TARGETS:
            path = DOCX_DIR / doc_name
            if not path.exists():
                print(f"Skip: {doc_name} not found")
                continue
            print(f"\nMeasuring {doc_name}...")
            all_results[doc_name] = measure_doc(word, path)
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(all_results, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUT}")

    print("\n=== Formula Residuals Summary ===")
    print(f"{'doc':<45s} {'tbl':>4} {'multi':>5} {'TE':>2} {'last_y':>7} {'lh':>5} {'body_y':>7} {'pred':>7} {'Δ':>6}")
    for doc, r in all_results.items():
        for t in r.get("tables", []):
            d = t["dominant_cell"]
            mp = "Y" if t["multi_page"] else "N"
            te = 1 if d["has_trailing_empty"] else 0
            by = t["first_body_after"]["y_pt"] if t["first_body_after"] else 0
            ly = d.get("last_page_last_y") or 0
            lh = d.get("line_height") or 0
            pred = t.get("formula_prediction") or 0
            resid = t.get("formula_residual") or 0
            print(f"{doc[:45]:<45s} {t['table_idx']:>4} {mp:>5} {te:>2} {ly:>7.2f} {lh:>5.2f} {by:>7.2f} {pred:>7.2f} {resid:>6.2f}")


if __name__ == "__main__":
    main()
