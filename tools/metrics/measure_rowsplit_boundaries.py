"""Measure Word's row-split boundaries across multiple docs + tables.

For each (doc, table_idx), find:
- Cell paragraph content (for wrap verification)
- Last line Y on old page (last continuation line before split)
- Continuation box bottom Y on new page (last cell line on new page)
- First body paragraph Y after the table on new page
- Continuation line count (how many lines on new page)

This builds the (continuation_lines, Word_first_body_y, cell_bottom_y)
dataset for cursor_y formula derivation.
"""
import json
from pathlib import Path
import win32com.client as w32


DOCX_DIR = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx")
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\rowsplit_boundaries.json")

# (doc_name, table_indices_to_check)
TARGETS = [
    ("d77a58485f16_20240705_resources_data_outline_08.docx", [5, 8]),
    ("d4d126dfe1d9_tokumei_08_01-3.docx", [4]),
    ("a1d6e4efa2e7_tokumei_08_01-4.docx", [5]),
    ("e3c545fac7a7_LOD_Handbook.docx", None),  # scan all
]


def measure_doc(word, docx_path: Path, table_indices=None):
    result = {"file": docx_path.name, "tables": []}
    doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
    try:
        result["page_count"] = doc.ComputeStatistics(2)
        if table_indices is None:
            # Scan for row-split tables (tables spanning multiple pages)
            table_indices = []
            for ti in range(1, doc.Tables.Count + 1):
                try:
                    tbl = doc.Tables(ti)
                    start_pg = tbl.Range.Information(3)
                    end_pg = doc.Range(tbl.Range.End - 1, tbl.Range.End).Information(3)
                    if end_pg > start_pg:
                        table_indices.append(ti)
                except Exception:
                    continue
            print(f"  Found {len(table_indices)} row-split tables: {table_indices}")

        for ti in table_indices:
            try:
                tbl = doc.Tables(ti)
            except Exception as e:
                print(f"  Table {ti}: skip ({e})")
                continue
            if tbl.Rows.Count != 1 or tbl.Columns.Count != 1:
                print(f"  Table {ti}: multi-row/col ({tbl.Rows.Count}x{tbl.Columns.Count}) -- skip for now")
                continue
            cell = tbl.Cell(1, 1)
            cell_range = cell.Range
            # Use start-position page specifically (Information(3) on whole range returns END page)
            start_r = doc.Range(cell_range.Start, cell_range.Start + 1)
            start_page = start_r.Information(3)

            # Find cell paragraphs and their line distribution
            cell_paras = []
            new_page_lines = []
            for pi in range(1, cell_range.Paragraphs.Count + 1):
                para = cell_range.Paragraphs(pi)
                pr = para.Range
                lines = []
                prev_y = None
                prev_page = None
                for off in range(pr.Start, pr.End):
                    r = doc.Range(off, off + 1)
                    try:
                        pg = r.Information(3)
                        y = r.Information(6)
                    except Exception:
                        continue
                    if prev_y is None or abs(y - prev_y) > 0.3 or pg != prev_page:
                        lines.append({"offset": off, "page": int(pg), "y_pt": round(y, 2)})
                        prev_y = y
                        prev_page = pg
                    # Capture new-page lines
                    if pg > start_page:
                        if not new_page_lines or abs(y - new_page_lines[-1]["y_pt"]) > 0.3 or pg != new_page_lines[-1]["page"]:
                            new_page_lines.append({"offset": off, "page": int(pg), "y_pt": round(y, 2)})
                cell_paras.append({
                    "p_idx": pi,
                    "chars": pr.End - pr.Start,
                    "preview": pr.Text[:30].replace("\r", "¶").replace("\x07", "⌂"),
                    "line_count": len(lines),
                })

            # Filter new_page_lines for actual new-page content only (dedupe by y)
            seen = set()
            filtered = []
            for ln in new_page_lines:
                key = (ln["page"], round(ln["y_pt"], 1))
                if key not in seen:
                    seen.add(key)
                    filtered.append(ln)
            # Sort by (page, y)
            filtered.sort(key=lambda x: (x["page"], x["y_pt"]))

            # First body paragraph after the cell
            cell_end = cell_range.End
            after_body = None
            for i, p in enumerate(doc.Paragraphs, start=1):
                pr = p.Range
                if pr.Start < cell_end:
                    continue
                try:
                    y = pr.Information(6)
                    pg = pr.Information(3)
                    in_table = bool(pr.Information(12))
                except Exception:
                    continue
                if in_table:
                    continue
                after_body = {
                    "idx": i,
                    "page": int(pg),
                    "y_pt": round(y, 2),
                    "text": pr.Text[:40].replace("\r", "¶").replace("\x07", "⌂"),
                }
                break

            t_info = {
                "table_idx": ti,
                "start_page": int(start_page),
                "cell_paras": cell_paras,
                "new_page_lines": filtered,
                "continuation_line_count": len(filtered),
                "first_body_after": after_body,
            }
            result["tables"].append(t_info)
            print(f"  Table {ti}: start_page={start_page} continuation={len(filtered)} lines first_body={after_body}")

    finally:
        doc.Close(SaveChanges=0)
    return result


def main():
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    all_results = {}
    try:
        for doc_name, targets in TARGETS:
            path = DOCX_DIR / doc_name
            if not path.exists():
                print(f"Skip: {doc_name} not found")
                continue
            print(f"\nMeasuring {doc_name}...")
            all_results[doc_name] = measure_doc(word, path, targets)
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(all_results, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUT}")

    # Summary table
    print("\n=== Summary ===")
    print(f"{'doc':40s} {'tbl':>4} {'cont':>5} {'first_y':>8}")
    for doc, r in all_results.items():
        for t in r.get("tables", []):
            fy = t["first_body_after"]["y_pt"] if t["first_body_after"] else 0
            print(f"{doc[:40]:40s} {t['table_idx']:>4} {t['continuation_line_count']:>5} {fy:>8.2f}")


if __name__ == "__main__":
    main()
