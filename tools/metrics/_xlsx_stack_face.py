# -*- coding: utf-8 -*-
"""Which face does Excel put a stacked character down through?

A stacked cell is drawn here through the vertical face — "@ＭＳ 明朝" turned a
quarter turn — because that face is what turns ー （ ） 「 」 and leaves
everything else upright, which is what Excel shows. The shapes agree; the
RASTER does not. A turned face is rasterised from outlines, and ＭＳ 明朝
carries embedded bitmaps for the sizes a sheet uses, so the upright face puts
down a different, crisper glyph than the turned one at the same pixel size.

So: character by character, is Excel's ink the UPRIGHT face's raster or the
TURNED face's? Both are drawn here with GDI at the same pixel size and
compared against Excel's own picture, bbox against bbox, pixel against pixel.

    python tools\\metrics\\_xlsx_stack_face.py
"""

from __future__ import annotations

import ctypes
import sys
import time
from ctypes import wintypes
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
SCRATCH = Path(r"C:\tmp\xlsx_stack_face")
GDI = ctypes.windll.gdi32

# One character per row: ideographs, both kana, the marks a vertical page turns,
# the ones it does not, full-width and half-width, and a few enclosed forms.
LETTERS = list(
    "相談政府統計あいウエ亜一二"
    "ー～〜｜（）「」【】［］｛｝〈〉《》"
    "、。・：；！？＝－＋／＼％＆＃＊＜＞￥"
    "０１ＡＢA1ｱｰ①⑧ⅠⅡ々〆㈱"
)
ARMS = [("ＭＳ 明朝", 11.0), ("ＭＳ ゴシック", 11.0), ("ＭＳ 明朝", 8.0)]
ROW_PT = 27.0  # 36px — room for one character of any size measured here


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
                ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
                ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD)]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 3)]


SIDE = 64


def raster(face: str, px: int, letter: str, turned: bool) -> np.ndarray:
    """The ink GDI puts down for one character, as a bool array of the box."""
    dc = GDI.CreateCompatibleDC(None)
    info = BITMAPINFO()
    head = info.bmiHeader
    head.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    head.biWidth, head.biHeight = SIDE, -SIDE
    head.biPlanes, head.biBitCount, head.biCompression = 1, 32, 0
    bits = ctypes.c_void_p()
    bitmap = GDI.CreateDIBSection(dc, ctypes.byref(info), 0, ctypes.byref(bits), None, 0)
    GDI.SelectObject(dc, bitmap)
    room = (ctypes.c_uint32 * (SIDE * SIDE)).from_address(bits.value)
    for at in range(SIDE * SIDE):
        room[at] = 0x00FFFFFF
    turn = 2700 if turned else 0
    name = f"@{face}" if turned else face
    font = GDI.CreateFontW(-px, 0, turn, turn, 400, 0, 0, 0, 1, 0, 0, 5, 0, name)
    GDI.SelectObject(dc, font)
    GDI.SetTextColor(dc, 0x000000)
    GDI.SetBkMode(dc, 1)
    GDI.SetTextAlign(dc, 0)  # TA_TOP | TA_LEFT
    held = ctypes.create_unicode_buffer(letter)
    GDI.TextOutW(dc, 20, 20, held, len(letter))
    seen = np.frombuffer(bytes(bytearray(room)), dtype=np.uint32).reshape(SIDE, SIDE)
    lit = (((seen & 0xFF) + ((seen >> 8) & 0xFF) + ((seen >> 16) & 0xFF)) // 3) < 140
    GDI.DeleteObject(font)
    GDI.DeleteObject(bitmap)
    GDI.DeleteDC(dc)
    return lit


def cropped(ink: np.ndarray) -> np.ndarray | None:
    rows, cols = np.where(ink)
    if len(rows) == 0:
        return None
    return ink[rows.min():rows.max() + 1, cols.min():cols.max() + 1]


def same(one: np.ndarray | None, two: np.ndarray | None) -> bool:
    if one is None or two is None:
        return one is two
    return one.shape == two.shape and bool((one == two).all())


def shot_of(face: str, size: float) -> tuple[np.ndarray, int, int] | None:
    """Excel's own picture of one stacked character per row."""
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "face.xlsx"
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:D{len(LETTERS) + 4}").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = 4.0
        for at, letter in enumerate(LETTERS, start=2):
            cell = sheet.Cells(at, 2)
            cell.Value = letter
            cell.Font.Name = face
            cell.Font.Size = size
            cell.Orientation = -4166          # xlVertical — stacked
            cell.VerticalAlignment = -4160    # xlTop
            cell.HorizontalAlignment = -4131  # xlLeft
            sheet.Rows(at).RowHeight = ROW_PT
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(1 + len(LETTERS), 2))
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(12):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(0.6)
            grabbed = ImageGrab.grabclipboard()
            if grabbed is not None:
                break
        else:
            print("  Excel would not hand over a picture")
            return None
        grabbed.save(SCRATCH / f"excel_{face}_{size}.png")
        across = round(sheet.Cells(2, 2).Width * 96 / 72)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    truth = np.asarray(Image.open(SCRATCH / f"excel_{face}_{size}.png").convert("L"))
    return truth < 140, across, round(ROW_PT * 96 / 72)


def main() -> int:
    for face, size in ARMS:
        px = round(size * 96 / 72)
        got = shot_of(face, size)
        if got is None:
            return 1
        truth, across, tall = got
        print(f"\n  {face} {size}pt — {px}px em, rows {tall}px, cell {across}px")
        print("  char  Excel bbox  upright  turned   verdict")
        tally: dict[str, list[str]] = {}
        for at, letter in enumerate(LETTERS):
            band = truth[at * tall:(at + 1) * tall, :across]
            theirs = cropped(band)
            plain = cropped(raster(face, px, letter, turned=False))
            turn = cropped(raster(face, px, letter, turned=True))
            shape = "-" if theirs is None else f"{theirs.shape[1]}x{theirs.shape[0]}"
            hit_plain = same(theirs, plain)
            hit_turn = same(theirs, turn)
            verdict = ("UPRIGHT" if hit_plain and not hit_turn else
                       "TURNED" if hit_turn and not hit_plain else
                       "both" if hit_plain else "neither")
            tally.setdefault(verdict, []).append(letter)
            print(f"   {letter}     {shape:>9}  {'yes' if hit_plain else '.':>7}"
                  f"  {'yes' if hit_turn else '.':>6}   {verdict}")
        print("  ---")
        for verdict, letters in sorted(tally.items()):
            print(f"  {verdict:>8} ({len(letters):>2}): {''.join(letters)}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
