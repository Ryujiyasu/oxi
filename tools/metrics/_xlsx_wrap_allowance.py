# -*- coding: utf-8 -*-
"""How much room beyond its text does a line keep before Excel wraps it?

The break positions are settled (SX28); what is not is the width the line is
measured against. This asks Excel directly and exactly: for a text of known
width, narrow the column a pixel at a time until the row grows a second line.
The narrowest column that still holds one line is the text's width plus the
allowance, so the allowance falls out by subtraction — no pictures, no
matching, just `Rows(1).Height` through COM.

    python tools\\metrics\\_xlsx_wrap_allowance.py
"""
import ctypes
import sys
from ctypes import wintypes

import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
GDI = ctypes.windll.gdi32
USER = ctypes.windll.user32

FACES = ["ＭＳ ゴシック", "ＭＳ Ｐゴシック", "ＭＳ 明朝", "游ゴシック",
         "Meiryo UI", "Calibri", "Arial"]
SIZES = [8.0, 9.0, 10.0, 11.0, 12.0, 14.0, 16.0, 18.0, 20.0, 24.0]
# Enough characters that the text is wide but the column never has to be
# wider than Excel allows.
SAMPLE = "あいうえおかきくけこさしすせそ"
LATIN = "abcdefghijklmnopqrstuvwxyz"
CASES = [(face, size, bold, (LATIN if face in ("Calibri", "Arial") else SAMPLE)[:max(3, int(70 / size))])
         for face in FACES for size in SIZES for bold in (False, True)]


class SIZE(ctypes.Structure):
    _fields_ = [("cx", wintypes.LONG), ("cy", wintypes.LONG)]


def run_width(face, points, bold, text):
    """Each character's own advance, added up — how Excel measures a line."""
    pixels = round(points * 96.0 / 72.0)
    dc = USER.GetDC(None)
    font = GDI.CreateFontW(-pixels, 0, 0, 0, 700 if bold else 400, 0, 0, 0,
                           1, 0, 0, 5, 0, face)
    old = GDI.SelectObject(dc, font)
    total = 0
    for letter in text:
        size = SIZE()
        GDI.GetTextExtentPoint32W(dc, letter, 1, ctypes.byref(size))
        total += size.cx
    GDI.SelectObject(dc, old)
    GDI.DeleteObject(font)
    USER.ReleaseDC(None, dc)
    return total


def main():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    print(f"{'face':<14}{'pt':>5}{'bold':>6}{'text px':>9}"
          f"{'narrowest column that holds one line':>38}{'allowance':>11}")
    try:
        book = excel.Workbooks.Add()
        sheet = book.Worksheets(1)
        for face, points, bold, text in CASES:
            cell = sheet.Cells(1, 1)
            sheet.Cells.Clear()
            cell.Value = text
            cell.Font.Name = face
            cell.Font.Size = points
            cell.Font.Bold = bold
            cell.WrapText = True
            sheet.Rows(1).RowHeight = 15  # so the height is free to grow
            sheet.Rows(1).AutoFit()

            def height_at(chars):
                sheet.Columns(1).ColumnWidth = chars
                sheet.Rows(1).AutoFit()
                return float(sheet.Rows(1).Height), int(round(sheet.Columns(1).Width / 0.75))

            # One line at its widest, then narrow until it is two.
            one, _ = height_at(80.0)
            wide, narrow = 80.0, 0.5
            for _ in range(40):
                middle = (wide + narrow) / 2
                held, _ = height_at(middle)
                if held > one + 0.01:
                    narrow = middle
                else:
                    wide = middle
            _, pixels = height_at(wide)
            width = run_width(face, points, bold, text)
            print(f"{face:<14}{points:>5.1f}{str(bold):>6}{width:>9}"
                  f"{pixels:>38}{pixels - width:>11}")
        book.Close(False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
