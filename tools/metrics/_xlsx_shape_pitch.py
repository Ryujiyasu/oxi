# -*- coding: utf-8 -*-
"""Which of the device's numbers is the pitch of a shape's lines?

`_xlsx_shape_text.py` measured what Excel does: 1.29em a line for the ＭＳ
faces, Calibri and Arial, 1.67em for 游ゴシック and Meiryo UI, and 1.94-2.09em
for メイリオ. That spread is the fonts' own, so the rule is a formula over
their metrics rather than a constant. This puts Excel's measured pitch beside
every candidate the device offers.

    python tools\\metrics\\_xlsx_shape_pitch.py
"""
import ctypes
import sys
from ctypes import wintypes

sys.stdout.reconfigure(encoding="utf-8")

GDI = ctypes.windll.gdi32
USER = ctypes.windll.user32


class TEXTMETRICW(ctypes.Structure):
    _fields_ = [("tmHeight", wintypes.LONG), ("tmAscent", wintypes.LONG),
                ("tmDescent", wintypes.LONG), ("tmInternalLeading", wintypes.LONG),
                ("tmExternalLeading", wintypes.LONG), ("tmAveCharWidth", wintypes.LONG),
                ("tmMaxCharWidth", wintypes.LONG), ("tmWeight", wintypes.LONG),
                ("tmOverhang", wintypes.LONG), ("tmDigitizedAspectX", wintypes.LONG),
                ("tmDigitizedAspectY", wintypes.LONG), ("tmFirstChar", wintypes.WCHAR),
                ("tmLastChar", wintypes.WCHAR), ("tmDefaultChar", wintypes.WCHAR),
                ("tmBreakChar", wintypes.WCHAR), ("tmItalic", wintypes.BYTE),
                ("tmUnderlined", wintypes.BYTE), ("tmStruckOut", wintypes.BYTE),
                ("tmPitchAndFamily", wintypes.BYTE), ("tmCharSet", wintypes.BYTE)]


class PANOSE(ctypes.Structure):
    _fields_ = [(name, ctypes.c_byte) for name in
                ("bFamilyType", "bSerifStyle", "bWeight", "bProportion", "bContrast",
                 "bStrokeVariation", "bArmStyle", "bLetterform", "bMidline", "bXHeight")]


class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class RECTL(ctypes.Structure):
    _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG),
                ("right", wintypes.LONG), ("bottom", wintypes.LONG)]


class OUTLINETEXTMETRICW(ctypes.Structure):
    _fields_ = [
        ("otmSize", wintypes.UINT), ("otmTextMetrics", TEXTMETRICW),
        ("otmFiller", ctypes.c_byte), ("otmPanoseNumber", PANOSE),
        ("otmfsSelection", wintypes.UINT), ("otmfsType", wintypes.UINT),
        ("otmsCharSlopeRise", wintypes.LONG), ("otmsCharSlopeRun", wintypes.LONG),
        ("otmItalicAngle", wintypes.LONG), ("otmEMSquare", wintypes.UINT),
        ("otmAscent", wintypes.LONG), ("otmDescent", wintypes.LONG),
        ("otmLineGap", wintypes.UINT), ("otmsCapEmHeight", wintypes.UINT),
        ("otmsXHeight", wintypes.UINT), ("otmrcFontBox", RECTL),
        ("otmMacAscent", wintypes.LONG), ("otmMacDescent", wintypes.LONG),
        ("otmMacLineGap", wintypes.UINT), ("otmusMinimumPPEM", wintypes.UINT),
        ("otmptSubscriptSize", POINT), ("otmptSubscriptOffset", POINT),
        ("otmptSuperscriptSize", POINT), ("otmptSuperscriptOffset", POINT),
        ("otmsStrikeoutSize", wintypes.UINT), ("otmsStrikeoutPosition", wintypes.LONG),
        ("otmsUnderscoreSize", wintypes.LONG), ("otmsUnderscorePosition", wintypes.LONG),
        ("otmpFamilyName", wintypes.LPCSTR), ("otmpFaceName", wintypes.LPCSTR),
        ("otmpStyleName", wintypes.LPCSTR), ("otmpFullName", wintypes.LPCSTR),
    ]


# Face, size, and the pitch Excel drew (from _xlsx_shape_text.py).
MEASURED = [
    ("メイリオ", 14.0, 39.0), ("メイリオ", 11.0, 28.5), ("メイリオ", 9.0, 23.5),
    ("ＭＳ Ｐゴシック", 14.0, 24.0), ("ＭＳ Ｐゴシック", 11.0, 19.0), ("ＭＳ Ｐゴシック", 8.0, 14.0),
    ("ＭＳ ゴシック", 11.0, 19.0), ("游ゴシック", 11.0, 24.5), ("游ゴシック", 14.0, 31.0),
    ("Meiryo UI", 11.0, 24.5), ("Calibri", 11.0, 19.0), ("Calibri", 18.0, 31.5),
    ("ＭＳ 明朝", 11.0, 19.0), ("Arial", 11.0, 19.0),
]


def metrics(face, points):
    dc = USER.GetDC(None)
    height = -round(points * 96 / 72)
    font = GDI.CreateFontW(height, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, face)
    held = GDI.SelectObject(dc, font)
    size = GDI.GetOutlineTextMetricsW(dc, 0, None)
    buffer = ctypes.create_string_buffer(size)
    GDI.GetOutlineTextMetricsW(dc, size, buffer)
    otm = ctypes.cast(buffer, ctypes.POINTER(OUTLINETEXTMETRICW)).contents
    held_metrics = (
        otm.otmTextMetrics.tmHeight,
        otm.otmTextMetrics.tmExternalLeading,
        otm.otmAscent,
        otm.otmDescent,
        otm.otmLineGap,
        otm.otmMacAscent,
        otm.otmMacDescent,
        otm.otmMacLineGap,
        otm.otmEMSquare,
    )
    GDI.SelectObject(dc, held)
    GDI.DeleteObject(font)
    USER.ReleaseDC(None, dc)
    return held_metrics


print(f"{'face':<16}{'pt':>5}{'Excel':>7}{'tmH':>5}{'ext':>5}{'asc':>5}{'desc':>6}"
      f"{'gap':>5}{'mac a/d/g':>14}{'1.2(a-d+g)':>12}{'a-d+g':>7}")
for face, points, drawn in MEASURED:
    (tall, ext, asc, desc, gap, mac_a, mac_d, mac_g, em) = metrics(face, points)
    line = asc - desc + gap
    print(f"{face:<16}{points:>5}{drawn:>7.1f}{tall:>5}{ext:>5}{asc:>5}{desc:>6}{gap:>5}"
          f"{f'{mac_a}/{mac_d}/{mac_g}':>14}{1.2 * line:>12.1f}{line:>7}")
