# -*- coding: utf-8 -*-
r"""What does Excel write into `<cols>` when a column is resized?

The editor is about to learn to write column widths, and a `<col>` element
carries more than a width: `customWidth`, `bestFit`, a style, `hidden`. Writing
the width and leaving the rest as they were is a guess, and a wrong one shows
up as a column that looks right here and wrong in Excel.

So this asks. Each arm sets a column up one way, saves, and prints the `<cols>`
element out of the file.

    python tools\metrics\_xlsx_col_width_xml.py
"""

from __future__ import annotations

import os
import re
import sys
import zipfile
from pathlib import Path

import win32com.client

SCRATCH = Path(r"C:\tmp\xlsx_col_width")


def cols_of(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        name = next(n for n in zf.namelist() if n.endswith("sheet1.xml"))
        xml = zf.read(name).decode("utf-8")
    found = re.search(r"<cols>.*?</cols>", xml, re.S)
    return found.group(0) if found else "(no <cols> at all)"


def build(what: str, arrange) -> None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    out = SCRATCH / f"{what}.xlsx"
    if out.exists():
        os.remove(out)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        for row in range(1, 4):
            for col in range(1, 6):
                sheet.Cells(row, col).Value = f"cell {row} {col}"
        arrange(sheet)
        book.SaveAs(str(out), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    print(f"  {what}")
    print(f"      {cols_of(out)}")
    # And what the same column reports back through the object model, so the
    # stored number can be read against the one a person types.
    print()


def main() -> int:
    build("nothing touched", lambda sheet: None)

    def one_width(sheet):
        sheet.Columns(2).ColumnWidth = 10
    build("column B set to 10", one_width)

    def autofit(sheet):
        sheet.Columns(2).AutoFit()
    build("column B autofitted", autofit)

    def autofit_then_drag(sheet):
        sheet.Columns(2).AutoFit()
        sheet.Columns(2).ColumnWidth = 20
    build("column B autofitted then set to 20", autofit_then_drag)

    def middle_of_a_run(sheet):
        sheet.Range("A:E").ColumnWidth = 12
        sheet.Columns(3).ColumnWidth = 30
    build("a run of five, then C widened", middle_of_a_run)

    def hidden_and_wide(sheet):
        sheet.Columns(2).ColumnWidth = 18
        sheet.Columns(2).Hidden = True
    build("column B widened then hidden", hidden_and_wide)

    def back_to_default(sheet):
        sheet.Columns(2).ColumnWidth = 25
        sheet.Columns(2).ColumnWidth = sheet.StandardWidth
    build("column B widened then put back", back_to_default)

    # The number a person types against the number the file stores.
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        print("  typed -> ColumnWidth -> Width in points")
        for typed in (1, 2, 5, 8.43, 10, 12.5, 20, 50):
            sheet.Columns(1).ColumnWidth = typed
            print(f"      {typed:>7} -> {sheet.Columns(1).ColumnWidth:>7} ->"
                  f" {sheet.Columns(1).Width:>8}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
