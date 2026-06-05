# -*- coding: utf-8 -*-
"""S498 cell-Y: COM-measure Word's per-(table,row) leading-cell first-paragraph content Y
(wdVerticalPositionRelativeToPage=6) with the R30 collapsed-start fix, + the row's top
border Y via Cell.Borders or the cell's range top. Also the cell's top margin (tcMar.top or
table default). cp932-safe: ASCII-only output to a file, no Japanese literals.

Usage: python _s498_cellY_com.py <docx> <out.json>
"""
import sys, os, json


def main():
    import win32com.client, pythoncom
    docx = os.path.abspath(sys.argv[1])
    out = sys.argv[2]
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    rows = []
    try:
        doc = word.Documents.Open(docx, ReadOnly=True)
        wdVert = 6  # wdVerticalPositionRelativeToPage
        wdPage = 3  # wdActiveEndPageNumber
        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(ti)
            nrows = tbl.Rows.Count
            for ri in range(1, nrows + 1):
                try:
                    row = tbl.Rows(ri)
                    cell = row.Cells(1)
                    rng = cell.Range
                    # collapsed start (R30) for content Y of the cell's first paragraph
                    s = rng.Start
                    cstart = doc.Range(s, s)
                    content_y = float(cstart.Information(wdVert))
                    page = int(cstart.Information(wdPage))
                    # row top from Cell.Range vs the row's vertical extent: use the cell's
                    # top edge approx = content_y - cell top inset. Word exposes Row.Height
                    # (may be wdRowHeightAuto=0). Capture HeightRule + Height.
                    try:
                        rheight = float(row.Height)
                    except Exception:
                        rheight = -1.0
                    try:
                        hrule = int(row.HeightRule)
                    except Exception:
                        hrule = -1
                    rows.append({
                        "tbl": ti, "row": ri, "page": page,
                        "content_y": round(content_y, 2),
                        "row_height": round(rheight, 2), "height_rule": hrule,
                    })
                except Exception as e:
                    rows.append({"tbl": ti, "row": ri, "error": str(e)[:80]})
        doc.Close(False)
    finally:
        word.Quit()
    json.dump({"rows": rows}, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("tables/rows measured:", len(rows), "-> wrote", out)


if __name__ == "__main__":
    main()
