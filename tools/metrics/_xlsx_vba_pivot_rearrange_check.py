# -*- coding: utf-8 -*-
"""Pivot tables a macro rearranged, saved by Oxi, opened in Excel.

The books are Excel's own pivot truth books (pipeline_data/xlsx_pivot), each
rearranged three ways by a VBA macro run in Oxi -- an item hidden, a field
moved to the other axis, a field added as a page field -- and saved (the
ignored wasm test `rearranged_pivot_books_for_excel` writes them). Here Excel
opens each saved book: it must open without repair and keep one pivot table,
and the cells Oxi wrote must equal what Excel shows after RefreshTable --
i.e. the saved definition is the rearranged one, not the book's old layout.

    OXI_PIVOT_REARRANGE_IN=pipeline_data/xlsx_pivot OXI_PIVOT_REARRANGE_OUT=<dir> \\
        cargo test -p oxidocs-wasm --features suite --lib rearranged_pivot_books -- --ignored
    python tools/metrics/_xlsx_vba_pivot_rearrange_check.py <dir>
"""
import sys
import time
from pathlib import Path

import json

import win32com.client

ROWS, COLS = 60, 20


def retry(fn, tries=50):
    for k in range(tries):
        try:
            return fn()
        except Exception as bad:
            if "-2147418111" not in str(bad) or k == tries - 1:
                raise
            time.sleep(0.2)


def norm(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, float):
        return round(v, 9)
    if v == "":
        return None
    return v


def oxi_cells(path):
    """The cells the macro left on the table's sheet (written beside the book)."""
    held = json.loads(path.with_suffix(".json").read_text(encoding="utf-8"))
    return held["sheet"], [[norm(v) for v in row[:COLS]] for row in held["cells"][:ROWS]]


def main():
    root = Path(sys.argv[1])
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    ok = total = 0
    try:
        for book in sorted(root.glob("*.xlsx")):
            total += 1
            try:
                wb = retry(lambda: xl.Workbooks.Open(str(book), 0, True, None, None, None, True, None, None, None, None, None, None, None, 0))
            except Exception as bad:
                print(f"  {book.stem:34s} EXCEL WOULD NOT OPEN {str(bad)[:100]}")
                continue
            sheets = [ws for ws in wb.Worksheets if retry(lambda: ws.PivotTables().Count) > 0]
            n = sum(retry(lambda: ws.PivotTables().Count) for ws in wb.Worksheets)
            if n != 1:
                print(f"  {book.stem:34s} PIVOTS={n}")
                wb.Close(False)
                continue
            sh = sheets[0]
            retry(lambda: sh.PivotTables(1).RefreshTable())
            excel = retry(lambda: [[norm(v) for v in row] for row in sh.Range(sh.Cells(1, 1), sh.Cells(ROWS, COLS)).Value])
            name = sh.Name
            wb.Close(False)
            sheet, oxi = oxi_cells(book)
            good = sheet == name and oxi == excel
            ok += good
            print(f"  {book.stem:34s} {'OK' if good else 'DIFF'}")
            if not good:
                for r, (a, b) in enumerate(zip(oxi, excel), 1):
                    if a != b:
                        print(f"      row {r} oxi   {[(i, v) for i, v in enumerate(a) if v is not None]}")
                        print(f"      row {r} excel {[(i, v) for i, v in enumerate(b) if v is not None]}")
    finally:
        xl.Quit()
    print(f"{ok}/{total} rearranged pivots open in Excel and match its refresh")


if __name__ == "__main__":
    main()
