# -*- coding: utf-8 -*-
"""Where is Excel's pen for the marks it lays on their side?

The standing characters are settled: Excel draws them through the upright face
with the pen on the row's own baseline (`_xlsx_stack_pen.py`). The turned ones —
（ ） ー 、 。 ～ ＝ and the rest of the class — are still drawn through the
vertical face, and there the pen is one pixel right and two down from ours at
11 point. This reads that pen against the device's own metrics over eleven
sizes and two faces, so the offset can be derived rather than tabulated.

The reference character in the first column gives Excel's upright pen, which is
the cell's geometry — gutters, centring, rounding — cancelled out.

    python tools\\metrics\\_xlsx_stack_turnpen.py
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
from _xlsx_stack_pen import line_box_px, offset  # noqa: E402

Image.MAX_IMAGE_PIXELS = None
SCRATCH = Path(r"C:\tmp\xlsx_stack_turnpen")
GDI = ctypes.windll.gdi32
TA_TOP, TA_BASELINE = 0, 24

STANDING = "相"
TURNED = ["（", "ー", "、"]
# Each column holds the reference character and then the mark, so both are
# placed by ONE cell's centring and it cancels out of the difference.
STACKS = [STANDING + STANDING] + [STANDING + mark for mark in TURNED]
SIZES = [6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0, 9.5, 10.0, 10.5, 11.0, 11.5, 12.0,
         12.5, 13.0, 13.5, 14.0, 15.0, 16.0, 17.0, 18.0, 19.0, 20.0, 22.0, 24.0]
FACES = ["ＭＳ 明朝"]
ROW_PT = 60.0
COLUMN = 9.0


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


def metrics(face: str, px: int, turned: bool) -> tuple[int, int, int]:
    dc = GDI.CreateCompatibleDC(None)
    turn = 2700 if turned else 0
    name = f"@{face}" if turned else face
    font = GDI.CreateFontW(-px, 0, turn, turn, 400, 0, 0, 0, 1, 0, 0, 5, 0, name)
    old = GDI.SelectObject(dc, font)
    told = TEXTMETRICW()
    GDI.GetTextMetricsW(dc, ctypes.byref(told))
    GDI.SelectObject(dc, old)
    GDI.DeleteObject(font)
    GDI.DeleteDC(dc)
    return told.tmAscent, told.tmDescent, told.tmHeight


def shot(face: str) -> tuple[np.ndarray, int, int] | None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    columns = len(STACKS)
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:G{len(SIZES) + 4}").Interior.Color = 0xFFFFFF
        for at in range(2, 2 + columns):
            sheet.Columns(at).ColumnWidth = COLUMN
            sheet.Columns(at).NumberFormat = "@"
        for at, size in enumerate(SIZES, start=2):
            for column, letter in enumerate(STACKS, start=2):
                cell = sheet.Cells(at, column)
                cell.Value = letter
                cell.Font.Name = face
                cell.Font.Size = size
                cell.Orientation = -4166          # xlVertical
                cell.VerticalAlignment = -4160    # xlTop
                cell.HorizontalAlignment = -4108  # xlCenter
            sheet.Rows(at).RowHeight = ROW_PT
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(1 + len(SIZES), 1 + columns))
        book.SaveAs(str(SCRATCH / "turnpen.xlsx"), FileFormat=51)
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


def ink(block: np.ndarray) -> tuple[int, int] | None:
    rows, cols = np.where(block)
    if len(rows) == 0:
        return None
    return int(cols.min()), int(rows.min())


def main() -> int:
    for face in FACES:
        got = shot(face)
        if got is None:
            print("  Excel would not hand over a picture")
            return 1
        truth, wide, tall = got
        print(f"\n  {face} — column {wide}px, row {tall}px, centred, sat on the top")
        print("  size  em  box  base | descU ascT descT pitch |"
              + "".join(f"   {letter}: dx  dy" for letter in TURNED))
        for at, size in enumerate(SIZES):
            px = round(size * 96 / 72)
            band = truth[at * tall:(at + 1) * tall]
            ascU, descU, _ = metrics(face, px, False)
            ascT, descT, tmH = metrics(face, px, True)
            pitch = line_box_px(size)
            told = ""
            for which, letter in enumerate(TURNED, start=1):
                cell = band[:, which * wide:(which + 1) * wide]
                # Each character is read inside ITS OWN line box. Reading the
                # cell whole takes the leftmost ink of BOTH characters for the
                # first one's, which is what made a turned mark look a pixel
                # further right than it is.
                head = ink(cell[:pitch])
                mark = ink(cell[pitch:2 * pitch + 5])
                guide = offset(face, px, STANDING, False, TA_TOP)
                turn_off = offset(face, px, letter, True, TA_BASELINE)
                if head is None or mark is None or turn_off is None:
                    told += "        -   -"
                    continue
                # Both characters stand in ONE cell, so its centring cancels.
                pen_up = (head[0] - guide[0], head[1] - guide[1])
                pen_turn = (mark[0] - turn_off[0], mark[1] + pitch - turn_off[1])
                told += f"   {pen_turn[0] - pen_up[0]:+4d} {pen_turn[1] - pen_up[1] - pitch:+3d}"
            print(f"  {size:>5} {px:>3} {line_box_px(size):>4} {'':>4} | {descU:>5} {ascT:>4}"
                  f" {descT:>5} {str(pitch):>5} |{told}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
