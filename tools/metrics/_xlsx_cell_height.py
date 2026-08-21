# -*- coding: utf-8 -*-
"""Is the height a CELL asks for the same as the sheet default for its font?

The row-height table is measured as `StandardHeight` after rewriting the
workbook's Normal font. fies_t2 says that is not the whole story: its rows of
Times New Roman 10 draw 20px where the table says 17, and rows of Century 9
draw 21 where it says 18 — both three pixels taller. A Latin face in a
Japanese Excel is paired with an East Asian one for the text it cannot show,
and the line may be measured against the pair.

So: for each face, the sheet default against what a row fits to with one
cell of that font, holding ASCII, holding Japanese, and holding nothing.
"""
import sys
import win32com.client


def main():
    faces = [
        ("Times New Roman", 10), ("Century", 9), ("Century", 11),
        ("Arial", 10), ("Calibri", 11), ("ＭＳ 明朝", 10),
        ("ＭＳ Ｐゴシック", 11), ("游ゴシック", 11), ("Terminal", 14),
    ]
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        normal = wb.Styles("Normal").Font
        print("%-18s %-5s %8s %8s %8s %8s" % (
            "face", "size", "default", "ascii", "japanese", "empty"))
        for face, size in faces:
            normal.Name = face
            normal.Size = size
            default = ws.StandardHeight / 0.75
            heights = []
            for value in ("Sample", "見本", None):
                ws.Cells.Clear()
                cell = ws.Range("A2")
                cell.Font.Name = face
                cell.Font.Size = size
                if value is not None:
                    cell.Value = value
                ws.Rows(2).AutoFit()
                heights.append(ws.Rows(2).Height / 0.75)
            print("%-18s %-5s %8.0f %8.0f %8.0f %8.0f%s" % (
                face, size, default, heights[0], heights[1], heights[2],
                "" if heights[0] == default else "   <- a cell asks for more"))
        wb.Close(False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
