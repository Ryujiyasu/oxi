"""
COM-measure Word's placement of the body paragraph AFTER_TABLE_BODY_PARA
following a row-split table. Compares to Oxi's cursor_y = page_top + overflow_on_next.
"""
import json
import os
from pathlib import Path
import win32com.client as w32


REPRO_DIR = Path(__file__).parent / "rowsplit_cursor_repro"
OUT_JSON = Path(__file__).parents[1] / "pipeline_data" / "rowsplit_cursor_measurements.json"


def measure(doc_path: Path):
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    result = {"file": doc_path.name}
    try:
        doc = word.Documents.Open(str(doc_path.resolve()), ReadOnly=True)
        try:
            page_count = doc.ComputeStatistics(2)  # wdStatisticPages = 2
            result["page_count"] = page_count

            # Find the body paragraph whose text starts with "AFTER_TABLE_BODY_PARA"
            paras = []
            for i, p in enumerate(doc.Paragraphs, start=1):
                text = p.Range.Text.rstrip("\x07\r\n ")
                in_cell = False
                try:
                    _ = p.Range.Information(12)  # wdWithInTable = 12
                    in_cell = bool(_)
                except Exception:
                    pass
                if text.startswith("AFTER_TABLE_BODY_PARA"):
                    y = p.Range.Information(6)   # wdVerticalPositionRelativeToPage
                    page = p.Range.Information(3)  # wdActiveEndPageNumber
                    paras.append({
                        "para_idx": i,
                        "text": text[:40],
                        "y_pt": y,
                        "page": page,
                        "in_cell": in_cell,
                    })
            result["after_table_paras"] = paras

            # Also measure last cell paragraph end (last table) final char position
            # to determine the overflow region.
            try:
                tbl = doc.Tables(doc.Tables.Count)
                cell = tbl.Cell(1, 1)
                cell_range = cell.Range
                # Each paragraph in the cell → its lines
                cell_paras = []
                for i, cp in enumerate(cell_range.Paragraphs, start=1):
                    r = cp.Range
                    lines = []
                    # Iterate through each character position to find line breaks
                    start = r.Start
                    end = r.End
                    prev_y = None
                    for offset in range(start, end):
                        char_range = doc.Range(offset, offset + 1)
                        try:
                            y = char_range.Information(6)
                            page = char_range.Information(3)
                            if prev_y is None or abs(y - prev_y) > 2:
                                lines.append({"offset": offset, "page": page, "y_pt": y})
                                prev_y = y
                        except Exception:
                            break
                    cell_paras.append({
                        "para_idx": i,
                        "chars": end - start,
                        "n_lines": len(lines),
                        "lines": lines,
                    })
                result["last_table"] = {
                    "num_paras": cell_range.Paragraphs.Count,
                    "paras": cell_paras,
                }
            except Exception as e:
                result["last_table_error"] = str(e)
        finally:
            doc.Close(SaveChanges=0)
    finally:
        word.Quit()
    return result


def main():
    OUT_JSON.parent.mkdir(exist_ok=True, parents=True)
    out = []
    for docx in sorted(REPRO_DIR.glob("*.docx")):
        print(f"Measuring {docx.name}...")
        result = measure(docx)
        out.append(result)
        print(f"  pages={result.get('page_count')}, after_table={result.get('after_table_paras')}")
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False)
    print(f"Wrote {OUT_JSON}")


if __name__ == "__main__":
    main()
