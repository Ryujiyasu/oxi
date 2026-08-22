# -*- coding: utf-8 -*-
"""What the device says about a face, next to what Excel does with it.

The cramped rows of `_xlsx_valign_pixels.py` — a line box taller than its row —
have Excel's ink starting at the row's top and ours two to five pixels down.
Before deciding what Excel raises the line by, ask GDI for the numbers a raise
could be made of: the ascent, the internal leading, and where the ink of the
probe's own string actually sits under the baseline.
"""
import ctypes
from ctypes import wintypes
import sys

gdi = ctypes.windll.gdi32
user = ctypes.windll.user32


class TEXTMETRICW(ctypes.Structure):
    _fields_ = [("tmHeight", wintypes.LONG), ("tmAscent", wintypes.LONG),
                ("tmDescent", wintypes.LONG), ("tmInternalLeading", wintypes.LONG),
                ("tmExternalLeading", wintypes.LONG), ("tmAveCharWidth", wintypes.LONG),
                ("tmMaxCharWidth", wintypes.LONG), ("tmWeight", wintypes.LONG),
                ("tmOverhang", wintypes.LONG), ("tmDigitizedAspectX", wintypes.LONG),
                ("tmDigitizedAspectY", wintypes.LONG), ("tmFirstChar", wintypes.WCHAR),
                ("tmLastChar", wintypes.WCHAR), ("tmDefaultChar", wintypes.WCHAR),
                ("tmBreakChar", wintypes.WCHAR), ("tmItalic", ctypes.c_byte),
                ("tmUnderlined", ctypes.c_byte), ("tmStruckOut", ctypes.c_byte),
                ("tmPitchAndFamily", ctypes.c_byte), ("tmCharSet", ctypes.c_byte)]


def face_metrics(face, points, bold):
    dc = user.GetDC(0)
    memory = gdi.CreateCompatibleDC(dc)
    height = -round(points * 96.0 / 72.0)
    font = gdi.CreateFontW(height, 0, 0, 0, 700 if bold else 400, 0, 0, 0,
                           1, 0, 0, 5, 0, face)
    gdi.SelectObject(memory, font)
    metric = TEXTMETRICW()
    gdi.GetTextMetricsW(memory, ctypes.byref(metric))
    gdi.DeleteObject(font)
    gdi.DeleteDC(memory)
    user.ReleaseDC(0, dc)
    return metric


CASES = [("ＭＳ Ｐゴシック", 11.0, False, 16, 18, 16, 0),
         ("ＭＳ Ｐゴシック", 14.0, False, 20, 22, 19, 0),
         ("ＭＳ ゴシック", 11.0, False, 16, None, None, 0),
         ("ＭＳ 明朝", 11.0, False, 16, None, None, 0),
         ("游ゴシック", 10.0, True, 16, None, None, 0),
         ("游ゴシック", 11.0, True, 18, 24, 18, 0),
         ("游ゴシック", 14.0, True, 23, None, None, 0),
         ("ＭＳ Ｐゴシック", 10.0, True, 14, None, None, 0),
         ("ＭＳ Ｐゴシック", 11.0, True, 16, 18, 16, 0),
         ("ＭＳ Ｐゴシック", 18.0, True, 24, None, None, 0),
         ("ＭＳ Ｐゴシック", 18.0, True, 26, None, None, 0),
         ("Calibri", 10.0, True, 14, 17, 13, 1)]

sys.stdout.reconfigure(encoding="utf-8")
print(f"{'face':<20}{'pt':>5}{'row':>5}  {'tmH':>4}{'asc':>4}{'desc':>5}{'ilead':>6}"
      f"{'asc-ilead':>10}")
for face, points, bold, row, _box, _base, _ in CASES:
    m = face_metrics(face, points, bold)
    print(f"{face + (' bold' if bold else ''):<20}{points:>5}{row:>5}  "
          f"{m.tmHeight:>4}{m.tmAscent:>4}{m.tmDescent:>5}{m.tmInternalLeading:>6}"
          f"{m.tmAscent - m.tmInternalLeading:>10}")
