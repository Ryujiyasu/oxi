# -*- coding: utf-8 -*-
"""How much taller is a row for holding a raised or lowered run?

`h2dee1989kre` row 5 fits to 27px where its font's line is 25, and every one
of its cells that holds `10³㎡` carries a `<vertAlign val="superscript"/>`
run. Clearing the row's contents drops it to 25; clearing only its formats
does not. So the raise is what asks for the extra, and this measures how much
across faces and sizes.

    python tools\\metrics\\_xlsx_raised_row.py
"""
import sys
from pathlib import Path

import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

REPO = Path(__file__).resolve().parents[2]
# Excel's own file, so the sweep runs against a real workbook's defaults.
BOOK = REPO / "tools" / "golden-test" / "documents" / "xlsx" / "ac821c1eea50_h2dee1989kre.xlsx"
FACES = ["游ゴシック", "メイリオ", "ＭＳ Ｐゴシック", "ＭＳ ゴシック", "Meiryo UI", "Calibri"]
SIZES = [8.0, 9.0, 10.0, 11.0, 12.0, 14.0, 16.0, 18.0, 20.0, 24.0]


def main():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(BOOK))
        sheet = book.Worksheets(1)
        spare = 200

        def fitted(face, size, raised):
            sheet.Rows(spare).ClearContents()
            sheet.Rows(spare).ClearFormats()
            cell = sheet.Cells(spare, 1)
            cell.Value = "10３㎡"
            cell.Font.Name = face
            cell.Font.Size = size
            if raised:
                cell.GetCharacters(3, 1).Font.Superscript = True
            sheet.Rows(spare).AutoFit()
            return sheet.Rows(spare).Height * 96 / 72

        print(f"{'face':<16}{'pt':>5}{'plain':>8}{'raised':>8}{'extra':>7}")
        for face in FACES:
            for size in SIZES:
                plain = fitted(face, size, False)
                raised = fitted(face, size, True)
                print(f"{face:<16}{size:>5}{plain:>8.1f}{raised:>8.1f}{raised - plain:>7.1f}")
        book.Close(SaveChanges=False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
