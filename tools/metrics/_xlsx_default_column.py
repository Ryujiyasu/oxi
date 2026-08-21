# -*- coding: utf-8 -*-
"""What width does Excel give a column the sheet says nothing about?

The renderer reads it as `trunc(8.43 x digit) + 5`, which matches every
workbook in the corpus but two: one whose standard font is Meiryo UI 12 and
one ＭＳ 明朝 14, where Excel draws 88px and the rule says 89. Both have a
10px digit, so the rule's shape is wrong somewhere the corpus rarely visits.

Sweep the standard font and read three things from Excel: the digit it
measures, the width in characters it reports, and the pixels it draws.
"""
import ctypes
import ctypes.wintypes as w
import json
import sys

import win32com.client

gdi = ctypes.windll.gdi32
user = ctypes.windll.user32


class SIZE(ctypes.Structure):
    _fields_ = [("cx", w.LONG), ("cy", w.LONG)]


def gdi_digit(face, points):
    hdc = user.GetDC(0)
    font = gdi.CreateFontW(-int(round(points * 96.0 / 72.0)), 0, 0, 0, 400,
                           0, 0, 0, 1, 0, 0, 0, 0, face)
    old = gdi.SelectObject(hdc, font)
    size = SIZE()
    body = ctypes.create_unicode_buffer("0")
    gdi.GetTextExtentPoint32W(hdc, body, 1, ctypes.byref(size))
    gdi.SelectObject(hdc, old)
    gdi.DeleteObject(font)
    user.ReleaseDC(0, hdc)
    return size.cx


FACES = [
    ("Calibri", 11), ("Calibri", 12), ("Calibri", 14),
    ("ＭＳ Ｐゴシック", 9), ("ＭＳ Ｐゴシック", 11), ("ＭＳ Ｐゴシック", 12), ("ＭＳ Ｐゴシック", 14),
    ("ＭＳ 明朝", 10), ("ＭＳ 明朝", 11), ("ＭＳ 明朝", 14),
    ("游ゴシック", 9), ("游ゴシック", 11), ("游ゴシック", 12), ("游ゴシック", 16),
    ("Meiryo UI", 10), ("Meiryo UI", 11), ("Meiryo UI", 12),
    ("メイリオ", 11), ("Yu Gothic UI", 12), ("Arial", 10), ("Arial", 12),
    ("Times New Roman", 10), ("Century", 11), ("Terminal", 14),
]


def main():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    rows = []
    try:
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        normal = wb.Styles("Normal").Font
        print("%-16s %-5s %5s %8s %7s | %s" % (
            "face", "size", "digit", "chars", "pixels", "candidate rules"))
        for face, size in FACES:
            normal.Name = face
            normal.Size = size
            chars = ws.StandardWidth
            pixels = round(ws.Columns(1).Width / 0.75)
            digit = gdi_digit(face, size)
            ours = int(8.43 * digit) + 5
            by_chars = int(chars * digit) + 5
            padded = int(((256 * chars + int(128 / digit)) / 256) * digit)
            rows.append({"face": face, "size": size, "digit": digit,
                         "chars": chars, "pixels": pixels})
            print("%-16s %-5s %5d %8s %7d | 8.43xd+5=%-4d chars*d+5=%-4d ooxml=%-4d%s" % (
                face, size, digit, chars, pixels, ours, by_chars, padded,
                "" if ours == pixels else "   <- the rule misses"))
        wb.Close(False)
    finally:
        excel.Quit()
    with open(r"pipeline_data\com_measurements\xlsx_default_column.json", "w",
              encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
