"""For each real docx, find every table cell paragraph that wraps across a
page break, and dump per-line (page, y_pt). Used to verify the line-level
row-split spec on ≥3 real docs.

Output: pipeline_data/row_split_real_doc_measurements.json
"""
import win32com.client
import json
import os
from pathlib import Path

REAL_DOCS = [
    "tools/golden-test/documents/docx/6514f214e482_tokumei_08_01-2.docx",
    "tools/golden-test/documents/docx/a1d6e4efa2e7_tokumei_08_01-4.docx",
    "tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx",
]

OUT_JSON = Path("pipeline_data") / "row_split_real_doc_measurements.json"


def find_splitting_cell_paras(word, docx_path: Path):
    """Scan all paragraphs. Return list of {para_start, para_end, table_idx, row, col,
       lines: [{offset, page, y}]} for paragraphs that (a) are inside a table cell
       and (b) span >1 page."""
    doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
    result = {
        "file": docx_path.name,
        "page_count": doc.ComputeStatistics(2),
        "splitting_paras": [],
    }
    try:
        n_paras = doc.Paragraphs.Count
        print(f"  scanning {n_paras} paras...")
        for pi in range(1, n_paras + 1):
            para = doc.Paragraphs(pi)
            pr = para.Range
            # Check if paragraph spans pages
            s_page = doc.Range(pr.Start, pr.Start + 1).Information(3) if pr.End > pr.Start else None
            e_page = doc.Range(max(pr.Start, pr.End - 1), pr.End).Information(3) if pr.End > pr.Start else None
            if s_page is None or e_page is None or s_page == e_page:
                continue
            # Check if in table
            in_tbl = pr.Information(12)  # wdWithInTable
            if not in_tbl:
                continue
            # Get table indices
            try:
                cell = pr.Cells(1)
                row_idx = cell.RowIndex
                col_idx = cell.ColumnIndex
                # Identify table via parent (ancestor) — use Tables.Count scan
                tbl_idx = None
                tables = doc.Tables
                for ti in range(1, tables.Count + 1):
                    t = tables(ti)
                    if t.Range.Start <= pr.Start <= t.Range.End:
                        tbl_idx = ti
                        break
            except Exception as ex:
                row_idx = col_idx = tbl_idx = -1
            # Enumerate per-line via Y jumps
            lines = []
            prev_y = None
            prev_page = None
            for off in range(pr.Start, pr.End):
                r = doc.Range(off, off + 1)
                pg = r.Information(3)
                y = r.Information(6)
                if prev_y is None or abs(y - prev_y) > 0.3 or pg != prev_page:
                    ch = r.Text[:1].replace('\r', '¶').replace('\n', '↵').replace('\t', '→')
                    if ch == '\x07':
                        ch = '⌂'
                    lines.append({"offset": off, "page": int(pg), "y_pt": round(y, 2), "char": ch})
                    prev_y = y
                    prev_page = pg
            pages_seen = sorted({ln["page"] for ln in lines})
            if len(pages_seen) <= 1:
                continue
            result["splitting_paras"].append({
                "para_idx": pi,
                "start": pr.Start,
                "end": pr.End,
                "table": tbl_idx,
                "row": row_idx,
                "col": col_idx,
                "line_count": len(lines),
                "pages": pages_seen,
                "lines": lines,
            })
        doc.Close(SaveChanges=False)
    except Exception as ex:
        try:
            doc.Close(SaveChanges=False)
        except Exception:
            pass
        result["error"] = str(ex)
    return result


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    out = []
    try:
        for relp in REAL_DOCS:
            p = Path(relp)
            if not p.exists():
                print(f"MISSING: {p}")
                continue
            print(f"\n=== {p.name} ===")
            r = find_splitting_cell_paras(word, p)
            print(f"  pages={r['page_count']} splitting_paras={len(r['splitting_paras'])}")
            for sp in r["splitting_paras"]:
                print(f"    para{sp['para_idx']:3d} tbl{sp['table']} row{sp['row']} col{sp['col']}  "
                      f"lines={sp['line_count']} pages={sp['pages']}")
                for i, ln in enumerate(sp["lines"]):
                    mark = ""
                    if i > 0 and ln["page"] != sp["lines"][i-1]["page"]:
                        mark = "  ** SPLIT **"
                    # Only print first, split, and last
                    if i <= 1 or i == len(sp["lines"]) - 1 or mark or i == 0 or (i > 0 and sp["lines"][i-1]["page"] != ln["page"]):
                        print(f"      line{i:2d} off={ln['offset']:6d} p{ln['page']} "
                              f"y={ln['y_pt']:7.2f} ch={ln['char']!r}{mark}")
            out.append(r)
    finally:
        word.Quit()

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nSaved → {OUT_JSON}")


if __name__ == "__main__":
    main()
