"""Where does Excel put an underline, and how far does it run?

`h2daa2023_dendeba_kmc` and its two siblings draw an underlined hyperlink
whose rule Excel puts one pixel lower than Oxi does and three pixels shorter.
167 underlined fonts across 79 workbooks carry the same risk, `data_A22`
among them.

GDI draws the underline itself when the font is made with the underline flag,
at whatever position the face declares. This asks Excel where it actually
puts one, and how far it runs against the text's own ink, over several faces
and sizes.

Run: python tools/metrics/_xlsx_underline.py
"""

from __future__ import annotations

import ctypes
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_underline")
WORDS = "Electronic tubes"
ARMS = [("ＭＳ Ｐゴシック", 11.0), ("ＭＳ Ｐゴシック", 14.0), ("Calibri", 11.0),
        ("メイリオ", 11.0), ("游ゴシック", 11.0), ("Arial", 11.0)]


class LOGFONT(ctypes.Structure):
    _fields_ = [("lfHeight", ctypes.c_long), ("lfWidth", ctypes.c_long),
        ("lfEscapement", ctypes.c_long), ("lfOrientation", ctypes.c_long),
        ("lfWeight", ctypes.c_long), ("lfItalic", ctypes.c_byte),
        ("lfUnderline", ctypes.c_byte), ("lfStrikeOut", ctypes.c_byte),
        ("lfCharSet", ctypes.c_byte), ("lfOutPrecision", ctypes.c_byte),
        ("lfClipPrecision", ctypes.c_byte), ("lfQuality", ctypes.c_byte),
        ("lfPitchAndFamily", ctypes.c_byte), ("lfFaceName", ctypes.c_wchar * 32)]


class TM(ctypes.Structure):
    _fields_ = [("tmHeight", ctypes.c_long), ("tmAscent", ctypes.c_long),
        ("tmDescent", ctypes.c_long), ("tmInternalLeading", ctypes.c_long),
        ("tmExternalLeading", ctypes.c_long), ("tmAveCharWidth", ctypes.c_long),
        ("tmMaxCharWidth", ctypes.c_long), ("tmWeight", ctypes.c_long),
        ("tmOverhang", ctypes.c_long), ("tmDigitizedAspectX", ctypes.c_long),
        ("tmDigitizedAspectY", ctypes.c_long), ("tmFirstChar", ctypes.c_wchar),
        ("tmLastChar", ctypes.c_wchar), ("tmDefaultChar", ctypes.c_wchar),
        ("tmBreakChar", ctypes.c_wchar), ("tmItalic", ctypes.c_byte),
        ("tmUnderlined", ctypes.c_byte), ("tmStruckOut", ctypes.c_byte),
        ("tmPitchAndFamily", ctypes.c_byte), ("tmCharSet", ctypes.c_byte)]


def by_gdi(face: str, points: float):
    """Where GDI draws the underline, and how wide the run is."""
    gdi = ctypes.windll.gdi32
    user = ctypes.windll.user32
    dc = user.GetDC(0)
    try:
        made = ctypes.windll.gdi32.CreateCompatibleDC(dc)
        wide = 900
        high = 80
        bitmap = gdi.CreateCompatibleBitmap(dc, wide, high)
        old_bitmap = gdi.SelectObject(made, bitmap)
        white = gdi.CreateSolidBrush(0x00FFFFFF)
        rect = ctypes.c_long * 4
        held = rect(0, 0, wide, high)
        user.FillRect(made, ctypes.byref(held), white)
        lf = LOGFONT()
        lf.lfHeight = -round(points * 96 / 72)
        lf.lfCharSet = 128
        lf.lfUnderline = 1
        lf.lfFaceName = face[:31]
        font = gdi.CreateFontIndirectW(ctypes.byref(lf))
        old = gdi.SelectObject(made, font)
        metrics = TM()
        gdi.GetTextMetricsW(made, ctypes.byref(metrics))
        gdi.SetBkMode(made, 1)
        gdi.TextOutW(made, 10, 10, WORDS, len(WORDS))
        # Read the bitmap back.
        class BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [("biSize", ctypes.c_uint32), ("biWidth", ctypes.c_long),
                ("biHeight", ctypes.c_long), ("biPlanes", ctypes.c_uint16),
                ("biBitCount", ctypes.c_uint16), ("biCompression", ctypes.c_uint32),
                ("biSizeImage", ctypes.c_uint32), ("biXPelsPerMeter", ctypes.c_long),
                ("biYPelsPerMeter", ctypes.c_long), ("biClrUsed", ctypes.c_uint32),
                ("biClrImportant", ctypes.c_uint32)]
        info = BITMAPINFOHEADER()
        info.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        info.biWidth = wide
        info.biHeight = -high
        info.biPlanes = 1
        info.biBitCount = 32
        buf = (ctypes.c_ubyte * (wide * high * 4))()
        gdi.GetDIBits(made, bitmap, 0, high, buf, ctypes.byref(info), 0)
        pixels = np.frombuffer(buf, dtype=np.uint8).reshape(high, wide, 4)[:, :, :3]
        gdi.SelectObject(made, old)
        gdi.DeleteObject(font)
        gdi.SelectObject(made, old_bitmap)
        gdi.DeleteObject(bitmap)
        gdi.DeleteObject(white)
        gdi.DeleteDC(made)
        ink = pixels.sum(axis=2) < 600
        rows = [(y, int(ink[y].sum())) for y in range(high) if ink[y].any()]
        widest = max(rows, key=lambda held: held[1]) if rows else (0, 0)
        lit = np.where(ink[widest[0]])[0]
        return {
            "rule y from the pen": widest[0] - 10,
            "rule x": (int(lit.min()) - 10, int(lit.max()) - 10) if len(lit) else None,
            "ascent": metrics.tmAscent,
        }
    finally:
        user.ReleaseDC(0, dc)


def picture(sheet):
    for _ in range(8):
        try:
            sheet.Activate()
            sheet.Range("A1:H12").CopyPicture(Appearance=1, Format=2)
        except Exception:
            time.sleep(0.6)
            continue
        time.sleep(0.4)
        held = ImageGrab.grabclipboard()
        if held is not None:
            return held
    return None


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:H12").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = 40.0
        sheet.Rows("3").RowHeight = 40.0
        cell = sheet.Range("B3")
        print("  face / size        Excel: rule row, x span        GDI: same, from its pen")
        for face, points in ARMS:
            cell.Value = WORDS
            cell.Font.Name = face
            cell.Font.Size = points
            cell.Font.Underline = 2      # xlUnderlineStyleSingle
            cell.HorizontalAlignment = -4131
            cell.VerticalAlignment = -4160
            held = picture(sheet)
            if held is None:
                print(f"  {face} {points}: no picture")
                continue
            held.save(SCRATCH / f"{face}_{points:.0f}.png")
            top = round(cell.Top * 96 / 72)
            left = round(cell.Left * 96 / 72)
            grey = np.asarray(held.convert("L"))[top:top + round(cell.Height * 96 / 72),
                                                 left:left + round(cell.Width * 96 / 72)]
            ink = grey < 140
            rows = [(y, int(ink[y].sum())) for y in range(ink.shape[0]) if ink[y].any()]
            if not rows:
                print(f"  {face} {points}: no ink")
                continue
            widest = max(rows, key=lambda held: held[1])
            lit = np.where(ink[widest[0]])[0]
            first = min(y for y, _ in rows)
            said = by_gdi(face, points)
            print(f"  {face:<14}{points:>5.0f}   rule at +{widest[0] - first:<3}"
                  f" x {int(lit.min())}..{int(lit.max())} ({widest[1]}px)"
                  f"    GDI +{said['rule y from the pen']} x {said['rule x']}"
                  f" ascent {said['ascent']}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
