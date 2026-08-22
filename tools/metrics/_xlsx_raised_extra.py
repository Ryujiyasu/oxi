# -*- coding: utf-8 -*-
"""How much room does a raised or lowered run ask for, per face and size?

A cell holding a `<vertAlign val="superscript"/>` run needs a taller line than
the same font without one — `h2dee1989kre` row 5 is 游ゴシック 11 whose line is
25px and Excel fits it to 27. This measures the difference on its own, in a
workbook whose blank row is short enough not to hide it.

    python tools\\metrics\\_xlsx_raised_extra.py
"""
import re
import sys
import zipfile
from pathlib import Path

import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

SCRATCH = Path(r"C:\tmp\xlsx_raised")
BOOK = SCRATCH / "raised.xlsx"
FACES = ["ＭＳ 明朝", "ＭＳ ゴシック", "ＭＳ Ｐゴシック", "ＭＳ Ｐ明朝",
         "游ゴシック", "メイリオ", "Meiryo UI", "Calibri", "游ゴシック Light"]
SIZES = [8.0, 9.0, 10.0, 10.5, 11.0, 12.0, 14.0, 16.0, 18.0, 20.0, 24.0]


def build():
    from openpyxl import Workbook

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.sheet_format.defaultRowHeight = 6.0
    for index in range(1, len(FACES) * len(SIZES) + 1):
        sheet.cell(row=index, column=1, value="10３㎡")
    book.save(BOOK)
    # A short default row, so a row's own content is what its height shows.
    with zipfile.ZipFile(BOOK) as source:
        parts = {item.filename: source.read(item.filename) for item in source.infolist()}
    sheet_xml = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    sheet_xml = re.sub(r'<sheetFormatPr[^>]*/>',
                       '<sheetFormatPr defaultRowHeight="6" customHeight="1"/>', sheet_xml)
    parts["xl/worksheets/sheet1.xml"] = sheet_xml.encode("utf-8")
    with zipfile.ZipFile(BOOK, "w", zipfile.ZIP_DEFLATED) as out:
        for name, body in parts.items():
            out.writestr(name, body)


def main():
    build()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(BOOK.resolve()))
        sheet = book.Worksheets(1)

        def fitted(index, face, size, raised):
            cell = sheet.Cells(index, 1)
            cell.Value = "10３㎡"
            cell.Font.Name = face
            cell.Font.Size = size
            if raised:
                cell.GetCharacters(3, 1).Font.Superscript = True
            sheet.Rows(index).AutoFit()
            return sheet.Rows(index).Height * 96 / 72

        print(f"{'face':<18}{'pt':>6}{'plain':>7}{'raised':>8}{'extra':>7}")
        row = 1
        for face in FACES:
            for size in SIZES:
                plain = fitted(row, face, size, False)
                raised = fitted(row, face, size, True)
                row += 1
                print(f"{face:<18}{size:>6}{plain:>7.0f}{raised:>8.0f}{raised - plain:>7.0f}")
        book.Close(SaveChanges=False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
