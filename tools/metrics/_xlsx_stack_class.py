# -*- coding: utf-8 -*-
"""Which marks does Excel lay on their side in a stacked cell?

`_xlsx_stack_pen.py` pinned where Excel's pen goes for a character it leaves
standing: on the row's own baseline, `top + baseline - ascent`, drawn through
the UPRIGHT face. That gives a decision procedure for every other character,
with no second law needed:

    predict the ink the upright face would leave at that pen.
    Excel's ink is that, to the pixel  ->  the character stands.
    it is not                          ->  the character is turned.

The pen is read from a reference character in the same column, so the cell's
own geometry — gutters, centring, rounding — cancels out.

    python tools\\metrics\\_xlsx_stack_class.py
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

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _xlsx_stack_face as FACE  # noqa: E402

Image.MAX_IMAGE_PIXELS = None
SCRATCH = Path(r"C:\tmp\xlsx_stack_class")
GDI = ctypes.windll.gdi32
TA_TOP = 0

REFERENCE = "相"
MARKS = (
    [chr(c) for c in range(0x3001, 0x3040)]            # CJK punctuation, kana marks
    + [chr(c) for c in range(0xFF01, 0xFF5F)]          # full-width forms
    + [chr(c) for c in range(0xFF61, 0xFFA0)]          # half-width katakana
    + list("‐‑‒–—―−…‥※〒◯○●■□△▲▽☆★♯℃％￥￡＄"
           "あアｱ亜一二三①⑧ⅠⅡ々〆")
)
ARMS = [("ＭＳ 明朝", 11.0), ("ＭＳ ゴシック", 11.0)]
ROW_PT = 27.0
COLUMN = 6.0


class SIZE(ctypes.Structure):
    _fields_ = [("cx", wintypes.LONG), ("cy", wintypes.LONG)]


def plain_ink(face: str, px: int, letter: str) -> tuple[int, int, int, int, int] | None:
    """The ink the upright face leaves, relative to the pen, and its advance."""
    pen, side = 40, 120
    dc = GDI.CreateCompatibleDC(None)
    info = FACE.BITMAPINFO()
    head = info.bmiHeader
    head.biSize = ctypes.sizeof(FACE.BITMAPINFOHEADER)
    head.biWidth, head.biHeight = side, -side
    head.biPlanes, head.biBitCount, head.biCompression = 1, 32, 0
    bits = ctypes.c_void_p()
    bitmap = GDI.CreateDIBSection(dc, ctypes.byref(info), 0, ctypes.byref(bits), None, 0)
    GDI.SelectObject(dc, bitmap)
    room = (ctypes.c_uint32 * (side * side)).from_address(bits.value)
    for at in range(side * side):
        room[at] = 0x00FFFFFF
    font = GDI.CreateFontW(-px, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, face)
    GDI.SelectObject(dc, font)
    GDI.SetTextColor(dc, 0x000000)
    GDI.SetBkMode(dc, 1)
    GDI.SetTextAlign(dc, TA_TOP)
    held = ctypes.create_unicode_buffer(letter)
    GDI.TextOutW(dc, pen, pen, held, len(letter))
    measured = SIZE()
    GDI.GetTextExtentPoint32W(dc, held, len(letter), ctypes.byref(measured))
    seen = np.frombuffer(bytes(bytearray(room)), dtype=np.uint32).reshape(side, side)
    lit = (((seen & 0xFF) + ((seen >> 8) & 0xFF) + ((seen >> 16) & 0xFF)) // 3) < 140
    GDI.DeleteObject(font)
    GDI.DeleteObject(bitmap)
    GDI.DeleteDC(dc)
    rows, cols = np.where(lit)
    if len(rows) == 0:
        return None
    return (int(cols.min()) - pen, int(rows.min()) - pen,
            int(cols.max() - cols.min() + 1), int(rows.max() - rows.min() + 1),
            int(measured.cx))


def shot(face: str, size: float, letters: list[str]) -> tuple[np.ndarray, int, int] | None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:D{len(letters) + 4}").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = COLUMN
        sheet.Columns("B").NumberFormat = "@"     # everything stays the text it is
        for at, letter in enumerate(letters, start=2):
            cell = sheet.Cells(at, 2)
            cell.Value = letter
            cell.Font.Name = face
            cell.Font.Size = size
            cell.Orientation = -4166          # xlVertical
            cell.VerticalAlignment = -4160    # xlTop
            cell.HorizontalAlignment = -4108  # xlCenter
            sheet.Rows(at).RowHeight = ROW_PT
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(1 + len(letters), 2))
        book.SaveAs(str(SCRATCH / "class.xlsx"), FileFormat=51)
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
            return None
        grabbed.save(SCRATCH / f"excel_{face}.png")
        wide = round(sheet.Cells(2, 2).Width * 96 / 72)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    truth = np.asarray(Image.open(SCRATCH / f"excel_{face}.png").convert("L")) < 140
    return truth, wide, round(ROW_PT * 96 / 72)


def inked(block: np.ndarray) -> tuple[int, int, int, int] | None:
    rows, cols = np.where(block)
    if len(rows) == 0:
        return None
    return (int(cols.min()), int(rows.min()),
            int(cols.max() - cols.min() + 1), int(rows.max() - rows.min() + 1))


def main() -> int:
    letters = [REFERENCE] + [m for m in MARKS if m.strip()]
    for face, size in ARMS:
        px = round(size * 96 / 72)
        got = shot(face, size, letters)
        if got is None:
            print("  Excel would not hand over a picture")
            return 1
        truth, wide, tall = got
        seen = inked(truth[0:tall, :wide])
        guide = plain_ink(face, px, REFERENCE)
        if seen is None or guide is None:
            print("  no reference to read the pen from")
            return 1
        pen = (seen[0] - guide[0], seen[1] - guide[1])
        print(f"\n  {face} {size}pt — {px}px, pen at {pen} in a {wide}px column")
        standing, turned, quiet = [], [], []
        for at, letter in enumerate(letters[1:], start=1):
            band = inked(truth[at * tall:(at + 1) * tall, :wide])
            plain = plain_ink(face, px, letter)
            if band is None or plain is None:
                quiet.append(letter)
                continue
            # A narrower character is centred on the full-width advance.
            shift = round((guide[4] - plain[4]) / 2)
            want = (pen[0] + shift + plain[0], pen[1] + plain[1], plain[2], plain[3])
            (standing if want == band else turned).append(letter)
        for name, group in (("standing", standing), ("TURNED", turned), ("no ink", quiet)):
            shown = "".join(group).encode("unicode_escape").decode("ascii")
            print(f"  {name} ({len(group)}): {shown}")
        print("  turned, one by one:")
        for letter in turned:
            band = inked(truth[(letters.index(letter)) * tall:(letters.index(letter) + 1) * tall, :wide])
            plain = plain_ink(face, px, letter)
            shift = round((guide[4] - plain[4]) / 2)
            want = (pen[0] + shift + plain[0], pen[1] + plain[1], plain[2], plain[3])
            print(f"    {letter.encode('unicode_escape').decode('ascii'):<10} Excel {band}  upright would be {want}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
