# -*- coding: utf-8 -*-
"""Is a row's height one font's line, or a box composed from several?

fies_t2's row 500 carries ＭＳ 明朝 8/10/11/12 and draws 19px — the tallest of
those fonts on its own. Row 226 carries the same and one cell of Century 9,
whose own line is 18px, and draws 21. A maximum cannot do that; a line box
built from the tallest ascent and the deepest descent can.

Measures each font's ascent and descent through GDI, and each font's own row
height through Excel, so the two can be held together.
"""
import ctypes
import ctypes.wintypes as w
import sys

import win32com.client

gdi = ctypes.windll.gdi32
user = ctypes.windll.user32


class TEXTMETRICW(ctypes.Structure):
    _fields_ = [
        ("tmHeight", w.LONG), ("tmAscent", w.LONG), ("tmDescent", w.LONG),
        ("tmInternalLeading", w.LONG), ("tmExternalLeading", w.LONG),
        ("tmAveCharWidth", w.LONG), ("tmMaxCharWidth", w.LONG),
        ("tmWeight", w.LONG), ("tmOverhang", w.LONG),
        ("tmDigitizedAspectX", w.LONG), ("tmDigitizedAspectY", w.LONG),
        ("tmFirstChar", w.WCHAR), ("tmLastChar", w.WCHAR),
        ("tmDefaultChar", w.WCHAR), ("tmBreakChar", w.WCHAR),
        ("tmItalic", ctypes.c_byte), ("tmUnderlined", ctypes.c_byte),
        ("tmStruckOut", ctypes.c_byte), ("tmPitchAndFamily", ctypes.c_byte),
        ("tmCharSet", ctypes.c_byte),
    ]


def metrics(face, points):
    hdc = user.GetDC(0)
    font = gdi.CreateFontW(-int(round(points * 96.0 / 72.0)), 0, 0, 0, 400,
                           0, 0, 0, 1, 0, 0, 0, 0, face)
    old = gdi.SelectObject(hdc, font)
    tm = TEXTMETRICW()
    gdi.GetTextMetricsW(hdc, ctypes.byref(tm))
    gdi.SelectObject(hdc, old)
    gdi.DeleteObject(font)
    user.ReleaseDC(0, hdc)
    return tm


FACES = [("ＭＳ 明朝", 8), ("ＭＳ 明朝", 9), ("ＭＳ 明朝", 10), ("ＭＳ 明朝", 11),
         ("ＭＳ 明朝", 12), ("Century", 9), ("Times New Roman", 10),
         ("Terminal", 14), ("游ゴシック", 11), ("Yu Gothic UI", 12)]


def main():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    rows = []
    try:
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        print("%-18s %-5s %6s %6s %6s %6s %6s" % (
            "face", "size", "excel", "asc", "desc", "a+d", "int"))
        for face, size in FACES:
            ws.Rows(3).Clear()
            ws.Rows(3).Font.Name = face
            ws.Rows(3).Font.Size = size
            ws.Rows(3).AutoFit()
            height = ws.Rows(3).Height / 0.75
            tm = metrics(face, size)
            rows.append((face, size, height, tm.tmAscent, tm.tmDescent))
            print("%-18s %-5s %6.0f %6d %6d %6d %6d" % (
                face, size, height, tm.tmAscent, tm.tmDescent,
                tm.tmAscent + tm.tmDescent, tm.tmInternalLeading))
        wb.Close(False)
    finally:
        excel.Quit()

    print("\nfies_t2's two rows, if the box is composed:")
    for label, present in [
        ("row 500 (ＭＳ 明朝 8/10/11/12)",
         [("ＭＳ 明朝", 8), ("ＭＳ 明朝", 10), ("ＭＳ 明朝", 11), ("ＭＳ 明朝", 12)]),
        ("row 226 (the same plus Century 9)",
         [("ＭＳ 明朝", 8), ("ＭＳ 明朝", 9), ("ＭＳ 明朝", 10), ("ＭＳ 明朝", 11),
          ("ＭＳ 明朝", 12), ("Century", 9)]),
        ("row 221 (plus Times New Roman 10)",
         [("ＭＳ 明朝", 8), ("ＭＳ 明朝", 10), ("ＭＳ 明朝", 11), ("ＭＳ 明朝", 12),
          ("Times New Roman", 10)]),
    ]:
        heights = [r for r in rows if (r[0], r[1]) in present]
        tallest = max(r[2] for r in heights)
        composed = max(r[3] for r in heights) + max(r[4] for r in heights)
        print("   %-36s tallest font %2.0fpx | composed %2dpx" % (
            label, tallest, composed))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
