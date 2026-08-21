# -*- coding: utf-8 -*-
"""What the device says about a font, beside what Excel does with it.

The glyphs now come out the size Excel draws them; what is left is where the
baseline goes. GDI's own numbers for the same font at the same pixel size are
the candidates, and this puts them next to the row height Excel gives a sheet
of that font, which the renderer already carries as a measured table.

    python tools\\metrics\\_xlsx_font_metrics.py
"""
import ctypes
import re
import sys
from ctypes import wintypes
from pathlib import Path

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

REPO = Path(__file__).resolve().parents[2]
TABLE = REPO / "tools" / "oxi-xlsx-renderer" / "src" / "row_defaults.rs"

FACES = ["ＭＳ ゴシック", "ＭＳ 明朝", "ＭＳ Ｐゴシック", "Meiryo UI", "Yu Gothic",
         "游ゴシック", "Calibri", "Arial"]
SIZES = [8.0, 9.0, 10.0, 10.5, 11.0, 12.0, 14.0, 16.0, 18.0]


def measured_rows():
    """The row height Excel gives a sheet of each font, as already measured."""
    held = {}
    for line in TABLE.read_text(encoding="utf-8").splitlines():
        found = re.search(r'\("([^"]+)",\s*(\d+),\s*(\d+)\)', line)
        if found:
            held[(found.group(1), int(found.group(2)) / 4.0)] = int(found.group(3))
    return held


def metrics(face, points, bold=False, italic=False):
    """The device's own account of a font at the pixel size Excel asks for."""
    pixels = round(points * 96.0 / 72.0)
    dc = USER.GetDC(None)
    font = GDI.CreateFontW(-pixels, 0, 0, 0, 700 if bold else 400, int(italic),
                           0, 0, 1, 0, 0, 4, 0, face)
    old = GDI.SelectObject(dc, font)
    tm = TEXTMETRICW()
    GDI.GetTextMetricsW(dc, ctypes.byref(tm))
    GDI.SelectObject(dc, old)
    GDI.DeleteObject(font)
    USER.ReleaseDC(None, dc)
    return pixels, tm


def ink_ascent(face, points, letter):
    """How far above the baseline the ink of this character starts.

    Drawn by the device itself at a baseline this knows, so the answer can be
    subtracted from ink measured in a picture to recover where the baseline
    that made it was.
    """
    pixels = round(points * 96.0 / 72.0)
    width, height, base = 200, 120, 80
    screen = USER.GetDC(None)
    dc = GDI.CreateCompatibleDC(screen)
    bitmap = GDI.CreateCompatibleBitmap(screen, width, height)
    GDI.SelectObject(dc, bitmap)
    GDI.PatBlt(dc, 0, 0, width, height, 0x00F00062)          # PATCOPY of white
    # Drawn without smoothing, because the picture this is compared against is
    # Excel's, and Excel's is the font's own bitmap: a smoothed edge would add
    # a faint row above the ink and read as a pixel more ascent than there is.
    font = GDI.CreateFontW(-pixels, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 3, 0, face)
    old = GDI.SelectObject(dc, font)
    GDI.SetBkMode(dc, 1)                                      # TRANSPARENT
    GDI.SetTextAlign(dc, 24)                                  # TA_BASELINE|TA_LEFT
    GDI.TextOutW(dc, 10, base, letter, len(letter))

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
                    ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
                    ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                    ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                    ("biClrImportant", wintypes.DWORD)]

    header = BITMAPINFOHEADER()
    header.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    header.biWidth = width
    header.biHeight = -height
    header.biPlanes = 1
    header.biBitCount = 32
    buffer = (ctypes.c_ubyte * (width * height * 4))()
    GDI.GetDIBits(dc, bitmap, 0, height, buffer, ctypes.byref(header), 0)
    GDI.SelectObject(dc, old)
    GDI.DeleteObject(font)
    GDI.DeleteObject(bitmap)
    GDI.DeleteDC(dc)
    USER.ReleaseDC(None, screen)

    top = None
    for y in range(height):
        row = buffer[y * width * 4:(y + 1) * width * 4:4]
        if any(value < 128 for value in row):
            top = y
            break
    return None if top is None else base - top


def main():
    rows = measured_rows()
    # Where Excel's own picture put the ink, from the glyph-size probe: one
    # character per row, top of the cell, rows of a known height.
    seen = excel_ink_tops()
    print(f"{'face':<14}{'pt':>6}{'em px':>7}{'ascent':>8}{'descent':>9}"
          f"{'height':>8}{'internal':>10}{'Excel row':>11}{'row−h':>7}"
          f"{'ink top':>9}{'baseline':>10}{'−ascent':>9}")
    for face in FACES:
        for points in SIZES:
            pixels, tm = metrics(face, points)
            row = rows.get((face, points))
            height = tm.tmHeight
            letter = "H" if face in ("Calibri", "Arial") else "亜"
            above = ink_ascent(face, points, letter)
            top = seen.get((face, points))
            baseline = None if top is None or above is None else top + above
            print(f"{face:<14}{points:>6.1f}{pixels:>7}{tm.tmAscent:>8}"
                  f"{tm.tmDescent:>9}{height:>8}{tm.tmInternalLeading:>10}"
                  f"{(row if row is not None else '-'):>11}"
                  f"{(row - height if row is not None else '-'):>7}"
                  f"{(top if top is not None else '-'):>9}"
                  f"{(baseline if baseline is not None else '-'):>10}"
                  f"{(baseline - tm.tmAscent if baseline is not None else '-'):>9}")


def excel_ink_tops():
    """Read the glyph-size probe's Excel picture: the ink top in each row."""
    import numpy as np
    from PIL import Image

    picture = Path(r"C:\tmp\xlsx_valign\glyph_size.excel.png")
    if not picture.exists():
        return {}
    image = np.asarray(Image.open(picture).convert("L"))
    faces = ["ＭＳ ゴシック", "ＭＳ 明朝", "ＭＳ Ｐゴシック", "Meiryo UI", "Yu Gothic",
             "Calibri"]
    sizes = [8.0, 9.0, 10.0, 10.5, 11.0, 12.0, 14.0, 16.0, 18.0]
    band, held, index = 40, {}, 0
    for face in faces:
        for points in sizes:
            strip = image[index * band:(index + 1) * band]
            dark = np.flatnonzero((strip < 128).any(axis=1))
            if dark.size:
                held[(face, points)] = int(dark[0])
            index += 1
    return held


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
