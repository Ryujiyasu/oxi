"""Where Excel puts an underline, measured from the baseline both engines share.

`_xlsx_underline.py` asked the question with two different rulers: the Excel
arm reported the rule against the cell's TOP INK ROW, the GDI arm against the
PEN it drew from. Those origins differ by the face's own ascent, so the 1px
the `dendeba_kmc` hyperlinks are out by could never be read off that table.

This asks it against a ruler both engines hold. `WORDS` has no descender, so
the last row of glyph ink IS the row above the baseline -- in a screenshot of
Excel and in a GDI bitmap alike. Every arm is therefore shot twice, once
plain and once underlined, and the answer is

    rule top - (last glyph row + 1)      i.e. rows below the baseline

which carries no origin at all. The x span is reported the same way, against
the run's own ink rather than against the cell.

Run: python tools/metrics/_xlsx_underline2.py
"""

from __future__ import annotations

import ctypes
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_underline2")
# No descender in any of these letters, so the bottom of the ink is the baseline.
WORDS = "Electronic tubes"
ARMS = (
    [("\uff2d\uff33 \uff30\u30b4\u30b7\u30c3\u30af", s)
     for s in (9.0, 10.0, 11.0, 12.0, 14.0, 16.0, 20.0)]
    + [("\uff2d\uff33 \u30b4\u30b7\u30c3\u30af", 11.0),
       ("\uff2d\uff33 \u660e\u671d", 11.0)]
    + [("\u30e1\u30a4\u30ea\u30aa", 11.0), ("\u6e38\u30b4\u30b7\u30c3\u30af", 11.0)]
    + [("Calibri", 11.0), ("Calibri", 14.0), ("Arial", 11.0),
       ("Times New Roman", 11.0)]
)
DARK = 160          # a ClearType glyph edge is lighter than this, its body is not
SOLID = 0.85        # a row this much of the run's width is the rule, not letters


class LOGFONT(ctypes.Structure):
    _fields_ = [("lfHeight", ctypes.c_long), ("lfWidth", ctypes.c_long),
        ("lfEscapement", ctypes.c_long), ("lfOrientation", ctypes.c_long),
        ("lfWeight", ctypes.c_long), ("lfItalic", ctypes.c_byte),
        ("lfUnderline", ctypes.c_byte), ("lfStrikeOut", ctypes.c_byte),
        ("lfCharSet", ctypes.c_byte), ("lfOutPrecision", ctypes.c_byte),
        ("lfClipPrecision", ctypes.c_byte), ("lfQuality", ctypes.c_byte),
        ("lfPitchAndFamily", ctypes.c_byte), ("lfFaceName", ctypes.c_wchar * 32)]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", ctypes.c_uint32), ("biWidth", ctypes.c_long),
        ("biHeight", ctypes.c_long), ("biPlanes", ctypes.c_uint16),
        ("biBitCount", ctypes.c_uint16), ("biCompression", ctypes.c_uint32),
        ("biSizeImage", ctypes.c_uint32), ("biXPelsPerMeter", ctypes.c_long),
        ("biYPelsPerMeter", ctypes.c_long), ("biClrUsed", ctypes.c_uint32),
        ("biClrImportant", ctypes.c_uint32)]


def read(ink):
    """The ink's own extent, and the rule if one is in it."""
    rows = [y for y in range(ink.shape[0]) if ink[y].any()]
    if not rows:
        return None
    lit = np.where(ink[rows].any(axis=0))[0]
    span = int(lit.max() - lit.min() + 1)
    counts = {y: int(ink[y].sum()) for y in rows}
    middle = (rows[0] + rows[-1]) / 2
    solid = [y for y in rows if counts[y] >= SOLID * span and y > middle]
    rule = None
    if solid:
        top = solid[0]
        thick = 1
        while top + thick in solid:
            thick += 1
        wide = np.where(ink[top])[0]
        rule = {"top": top, "thick": thick,
                "x": (int(wide.min()), int(wide.max()))}
    return {"top": rows[0], "bottom": rows[-1],
            "x": (int(lit.min()), int(lit.max())),
            "rule": rule, "counts": counts}


def by_gdi(face, points, underline, charset=1):
    """The same picture, drawn by GDI from a pen at (10, 10).

    `charset` mirrors the renderer, which asks for DEFAULT_CHARSET (1).
    Asking a Latin face for SHIFT-JIS (128) can make GDI realize a
    different face entirely, so the realized name is read back and
    reported -- a probe that measured a substituted face would otherwise
    look like a finding.
    """
    gdi = ctypes.windll.gdi32
    user = ctypes.windll.user32
    screen = user.GetDC(0)
    try:
        dc = gdi.CreateCompatibleDC(screen)
        wide, high = 900, 90
        bitmap = gdi.CreateCompatibleBitmap(screen, wide, high)
        old_bitmap = gdi.SelectObject(dc, bitmap)
        white = gdi.CreateSolidBrush(0x00FFFFFF)
        rect = (ctypes.c_long * 4)(0, 0, wide, high)
        user.FillRect(dc, ctypes.byref(rect), white)
        lf = LOGFONT()
        lf.lfHeight = -round(points * 96 / 72)
        lf.lfCharSet = charset
        lf.lfUnderline = 1 if underline else 0
        lf.lfQuality = 5          # CLEARTYPE_QUALITY, the renderer's own
        lf.lfFaceName = face[:31]
        font = gdi.CreateFontIndirectW(ctypes.byref(lf))
        old = gdi.SelectObject(dc, font)
        gdi.SetBkMode(dc, 1)
        name = ctypes.create_unicode_buffer(64)
        gdi.GetTextFaceW(dc, 64, name)
        gdi.TextOutW(dc, 10, 10, WORDS, len(WORDS))
        info = BITMAPINFOHEADER()
        info.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        info.biWidth, info.biHeight = wide, -high
        info.biPlanes, info.biBitCount = 1, 32
        buf = (ctypes.c_ubyte * (wide * high * 4))()
        gdi.GetDIBits(dc, bitmap, 0, high, buf, ctypes.byref(info), 0)
        pixels = np.frombuffer(buf, dtype=np.uint8).reshape(high, wide, 4)[:, :, :3]
        gdi.SelectObject(dc, old)
        gdi.DeleteObject(font)
        gdi.SelectObject(dc, old_bitmap)
        gdi.DeleteObject(bitmap)
        gdi.DeleteObject(white)
        gdi.DeleteDC(dc)
        held = read(pixels.mean(axis=2) < DARK)
        if held is not None:
            held["face"] = name.value
        return held
    finally:
        user.ReleaseDC(0, screen)


def picture(sheet, want):
    """A screen bitmap of the sheet, or None.

    The clipboard belongs to the whole machine and another session may be
    copying into it, so the picture is only accepted at the size this sheet
    asked for.
    """
    for _ in range(8):
        try:
            sheet.Activate()
            sheet.Range("A1:H12").CopyPicture(Appearance=1, Format=2)
        except Exception:
            time.sleep(0.6)
            continue
        time.sleep(0.4)
        held = ImageGrab.grabclipboard()
        if held is None:
            continue
        if abs(held.width - want[0]) > 4 or abs(held.height - want[1]) > 4:
            print("    (clipboard held {}x{}, wanted {}x{} -- reshooting)".format(
                held.width, held.height, want[0], want[1]))
            time.sleep(0.5)
            continue
        return held
    return None


def by_excel(sheet, cell, underline, want):
    cell.Font.Underline = 2 if underline else -4142
    held = picture(sheet, want)
    if held is None:
        return None
    top = round(cell.Top * 96 / 72)
    left = round(cell.Left * 96 / 72)
    grey = np.asarray(held.convert("L"))[
        top:top + round(cell.Height * 96 / 72),
        left:left + round(cell.Width * 96 / 72)]
    stem = "{}_{:.0f}{}.png".format(cell.Font.Name, cell.Font.Size,
                                    "_u" if underline else "")
    held.save(SCRATCH / stem.replace(" ", "_"))
    return read(grey < DARK)


def main():
    SCRATCH.mkdir(parents=True, exist_ok=True)
    # A separate instance: another session may own the Excel already running,
    # and Quit() on a shared one would take its measurements down with it.
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:H12").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = 40.0
        sheet.Rows("3").RowHeight = 40.0
        cell = sheet.Range("B3")
        cell.Value = WORDS
        cell.HorizontalAlignment = -4131      # xlLeft
        cell.VerticalAlignment = -4160        # xlTop
        want = (round(sheet.Range("A1:H12").Width * 96 / 72),
                round(sheet.Range("A1:H12").Height * 96 / 72))
        print("  ruler = the baseline; rule rows are counted below it."
              "  picture wanted {}x{}\n".format(want[0], want[1]))
        print("  {:<18}{:>4}   {:>17} {:>5}   {:>15} {:>5}   {:>3}"
              "   {:>6} {:>6}  GDI realized".format(
                  "face", "pt", "Excel below base", "thick",
                  "GDI below base", "thick", "d", "xl w", "gdi w"))
        for face, points in ARMS:
            cell.Font.Name = face
            cell.Font.Size = points
            plain = by_excel(sheet, cell, False, want)
            under = by_excel(sheet, cell, True, want)
            gplain = by_gdi(face, points, False)
            gunder = by_gdi(face, points, True)
            if not (plain and under and gplain and gunder
                    and under["rule"] and gunder["rule"]):
                print("  {:<18}{:>4.0f}   (no reading)".format(face, points))
                continue
            # The baseline is the row after the last row of a descender-free run.
            xl = under["rule"]["top"] - (plain["bottom"] + 1)
            gd = gunder["rule"]["top"] - (gplain["bottom"] + 1)
            xl_x = (under["rule"]["x"][0] - plain["x"][0],
                    under["rule"]["x"][1] - plain["x"][1])
            gd_x = (gunder["rule"]["x"][0] - gplain["x"][0],
                    gunder["rule"]["x"][1] - gplain["x"][1])
            xl_w = plain["x"][1] - plain["x"][0] + 1
            gd_w = gplain["x"][1] - gplain["x"][0] + 1
            same = "" if gplain["face"] == face else "  <-- NOT the face asked for"
            print("  {:<18}{:>4.0f}   {:>17} {:>5}   {:>15} {:>5}   {:>+3}"
                  "   {:>6} {:>6}  {}{}".format(
                      face, points, xl, under["rule"]["thick"],
                      gd, gunder["rule"]["thick"], xl - gd,
                      xl_w, gd_w, gplain["face"], same))
        print("\n  (Excel below base) - (GDI below base) is what Oxi is out by,"
              "\n  because Oxi lets GDI draw the rule from the face's own metric.")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
