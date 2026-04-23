"""Extended row-split boundary scan: handles multi-row/col tables.

For each table (regardless of structure), find the ROW whose content spans
multiple pages (= the split row). Within that row, find the cell with the
most continuation content on the new page. Treat that cell's
continuation as the row-split measurement target.

Also: first non-table paragraph y after the entire table.
"""
import json
import sys
from pathlib import Path
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx")
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\rowsplit_boundaries_v2.json")

TARGETS = [
    "d77a58485f16_20240705_resources_data_outline_08.docx",
    "d4d126dfe1d9_tokumei_08_01-3.docx",
    "a1d6e4efa2e7_tokumei_08_01-4.docx",
    "e3c545fac7a7_LOD_Handbook.docx",
]


def info_start_page(doc, rng):
    """Page of the start offset of a range."""
    try:
        return doc.Range(rng.Start, rng.Start + 1).Information(3)
    except Exception:
        return None


def info_end_page(doc, rng):
    """Page of the last char of a range."""
    try:
        return doc.Range(rng.End - 1, rng.End).Information(3)
    except Exception:
        return None


def measure_cell_continuation(doc, cell, row_start_page):
    """Count unique y positions on pages > row_start_page within a cell.
    Returns (continuation_line_count, first_new_page_y, last_new_page_y).
    """
    cell_range = cell.Range
    seen_ys = {}  # (page, rounded_y) -> first offset
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
        return 0, None, None
    # Filter outliers (y far from others - likely end-of-cell artifact)
    ys_sorted = sorted(seen_ys.keys())
    all_ys = [y for _, y in ys_sorted]
    # Identify cluster on first continuation page
    first_pg = ys_sorted[0][0]
    page_ys = [y for p, y in ys_sorted if p == first_pg]
    # Remove outliers: if a y is >100pt gap from previous, likely artifact
    filtered = [page_ys[0]]
    for y in page_ys[1:]:
        if y - filtered[-1] < 50:  # reasonable line spacing
            filtered.append(y)
        else:
            break  # stop at first big gap
    return len(filtered), filtered[0], filtered[-1]


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

            # Find split row: row where start page != end page
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
                continue  # no row spans pages

            # For each cell in split row, measure continuation
            cell_data = []
            for ci in range(1, cols + 1):
                try:
                    cell = split_row.Cells(ci)
                    cont_count, first_y, last_y = measure_cell_continuation(doc, cell, row_start_page)
                    if cont_count > 0:
                        cell_data.append({
                            "col": ci,
                            "continuation_line_count": cont_count,
                            "first_new_page_y": first_y,
                            "last_new_page_y": last_y,
                            "paragraphs": cell.Range.Paragraphs.Count,
                        })
                except Exception:
                    continue

            if not cell_data:
                continue

            # Get first non-table paragraph y after the entire table
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

            # Pick the cell with the most continuation content as the
            # dominant cell for formula derivation
            dominant_cell = max(cell_data, key=lambda c: c["continuation_line_count"])

            t_info = {
                "table_idx": ti,
                "rows": rows,
                "cols": cols,
                "split_row": split_row_idx,
                "row_start_page": row_start_page,
                "row_end_page": row_end_page,
                "cells_with_continuation": cell_data,
                "dominant_cell": dominant_cell,
                "first_body_after": after_body,
            }
            result["tables"].append(t_info)
            print(f"  Table {ti} ({rows}x{cols}): split_row={split_row_idx}, dominant cont={dominant_cell['continuation_line_count']}, first_body_y={after_body['y_pt'] if after_body else '?'}")
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

    # Summary
    print("\n=== Row-split Summary ===")
    print(f"{'doc':<45s} {'tbl':>4} {'rxc':>6} {'row':>4} {'cont':>5} {'first_y':>8}")
    for doc, r in all_results.items():
        for t in r.get("tables", []):
            fy = t["first_body_after"]["y_pt"] if t["first_body_after"] else 0
            rxc = f"{t['rows']}x{t['cols']}"
            print(f"{doc[:45]:<45s} {t['table_idx']:>4} {rxc:>6} {t['split_row']:>4} {t['dominant_cell']['continuation_line_count']:>5} {fy:>8.2f}")


if __name__ == "__main__":
    main()
