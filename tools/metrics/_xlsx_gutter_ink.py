# -*- coding: utf-8 -*-
r"""Where does a cell's first glyph actually start, across the column's edge?

`gutters()` was derived by narrowing a column until the text wrapped, which
measures the SUM of the two gutters — how much room the line is broken in. The
split between them (left keeps one more than the right) does not follow from
that measurement, and the floor workbook says the left one is short: its title
cell, Yu Gothic UI 22pt bold, starts a pixel left of Excel's, and every glyph
after it is identical once shifted.

So this asks the left gutter on its own, as ink. One cell an arm, left
aligned, holding a single `0` — a digit whose side bearing the device will
state, so the gutter is the ink's start less the column's edge less that
bearing. The right-aligned twin beside it reads the other gutter the same way.

    python tools\metrics\_xlsx_gutter_ink.py
    python tools\metrics\_xlsx_gutter_ink.py --reuse
"""

from __future__ import annotations

import argparse
import ctypes
import os
import subprocess
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SCRATCH = Path(r"C:\tmp\xlsx_gutter_ink_dense"
               if "--dense" in sys.argv
               else r"C:\tmp\xlsx_gutter_ink")
WORDS = "0"
ROW_PT = 30.0
COLUMN = 14.0
ARMS = []
if "--dense" in sys.argv:
    # Every whole point of the face the floor workbook wears, both weights:
    # the sparse sweep found two misses in it and cannot say where the step
    # that produces them belongs.
    for _face in ("Yu Gothic UI", "メイリオ"):
        for _step in range(8, 37):
            ARMS.append((_face, float(_step), False))
            ARMS.append((_face, float(_step), True))
else:
    for _face in ("ＭＳ ゴシック", "ＭＳ Ｐゴシック",
                  "游ゴシック", "Yu Gothic UI",
                  "メイリオ", "Calibri", "Arial", "ＭＳ 明朝"):
        # 5 and 6 point were never swept: the rule's `floor((digit - 5) / 4)`
        # goes negative there and is clamped, and `barrier_free` sets four of
        # its fonts at 6pt.
        for _size in (5.0, 6.0, 7.0, 8.0, 11.0, 14.0, 18.0, 22.0, 26.0, 36.0):
            ARMS.append((_face, _size, False))
            ARMS.append((_face, _size, True))


class ABC(ctypes.Structure):
    _fields_ = [("abcA", ctypes.c_int), ("abcB", ctypes.c_uint), ("abcC", ctypes.c_int)]


def bearing(face: str, points: float, bold: bool) -> tuple[int, int, int]:
    """The device's own A, B and C for `0` at this size."""
    gdi, user = ctypes.windll.gdi32, ctypes.windll.user32
    screen = user.GetDC(0)
    dc = gdi.CreateCompatibleDC(screen)
    font = gdi.CreateFontW(-round(points * 96 / 72), 0, 0, 0, 700 if bold else 400,
                           0, 0, 0, 1, 0, 0, 5, 0, face)
    held = gdi.SelectObject(dc, font)
    box = ABC()
    gdi.GetCharABCWidthsW(dc, ord(WORDS), ord(WORDS), ctypes.byref(box))
    gdi.SelectObject(dc, held)
    gdi.DeleteObject(font)
    gdi.DeleteDC(dc)
    user.ReleaseDC(0, screen)
    return box.abcA, box.abcB, box.abcC


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:E{len(ARMS) + 4}").Interior.Color = 0xFFFFFF
        sheet.Columns(2).ColumnWidth = COLUMN
        sheet.Columns(3).ColumnWidth = COLUMN
        for at, (face, points, bold) in enumerate(ARMS, start=2):
            for column, how in ((2, -4131), (3, -4152)):   # left, right
                cell = sheet.Cells(at, column)
                # A number, not text: a text number wears the green "stored as
                # text" corner, which is ink of its own at the cell's start.
                cell.Value = 0
                cell.NumberFormat = "0"
                cell.Font.Name = face
                cell.Font.Size = points
                cell.Font.Bold = bold
                cell.HorizontalAlignment = how
                cell.VerticalAlignment = -4160          # top, so rows cannot mix
            sheet.Rows(at).RowHeight = ROW_PT
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(2, 2), sheet.Cells(1 + len(ARMS), 3)).CopyPicture(
                    Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.8)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def ours(made: Path) -> tuple[np.ndarray, dict[int, tuple[int, int]], int]:
    told = subprocess.run([str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
                          env={"OXI_XLSX_DUMP_COLUMNS": "1", "OXI_XLSX_DUMP_ROWS": "1",
                               **os.environ},
                          capture_output=True, text=True, encoding="utf-8")
    columns, at, top, down = {}, 0, 0, 0
    for line in (told.stdout or "").splitlines() + (told.stderr or "").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "column":
            columns[int(parts[1])] = (at, at + int(parts[3]))
            at += int(parts[3])
        if len(parts) == 4 and parts[0] == "row":
            if int(parts[1]) == 2:
                top = down
            down += int(parts[3])
    return np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140, columns, top


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    parser.add_argument("--dense", action="store_true",
                        help="every whole point of one face, both weights")
    args = parser.parse_args()
    made = SCRATCH / "gutter.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    mine, columns, top = ours(made)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    tall = round(ROW_PT * 96 / 72)
    wide = columns[1][1] - columns[1][0]
    print(f"  a single {WORDS!r} in a {wide}px column, top aligned")
    print("  face             size  bold  digit  A  |  left: Excel Oxi  |  right: Excel Oxi")
    for at, (face, points, bold) in enumerate(ARMS):
        lead, _body, tail = bearing(face, points, bold)
        told = []
        for column in (0, 1):
            band = truth[at * tall + 1:(at + 1) * tall - 1,
                         column * wide + 1:(column + 1) * wide - 1]
            lit = np.where(band.any(axis=0))[0]
            theirs = (int(lit.min()) + 1, int(lit.max()) + 1) if len(lit) else None
            left, right = columns[column + 1]
            band = mine[top + at * tall + 1:top + (at + 1) * tall - 1, left + 1:right - 1]
            lit = np.where(band.any(axis=0))[0]
            ours_at = (int(lit.min()) + 1, int(lit.max()) + 1) if len(lit) else None
            told.append((theirs, ours_at))
        if any(one is None for pair in told for one in pair):
            print(f"  {face:<16}{points:>5}  {bold!s:>5}   nothing to read")
            continue
        digit = _body + lead + tail
        # Against the left edge the gutter is the ink's start less the glyph's
        # own lead; against the right it is the column's edge less the ink's
        # end less the glyph's tail.
        print(f"  {face:<16}{points:>5}  {bold!s:>5}  {digit:>5} {lead:>2}  |"
              f"  {told[0][0][0] - lead:>5} {told[0][1][0] - lead:>4}"
              f"  {'' if told[0][0][0] == told[0][1][0] else '<<':<3}|"
              f"  {wide - told[1][0][1] - tail:>5} {wide - told[1][1][1] - tail:>4}"
              f"  {'' if told[1][0][1] == told[1][1][1] else '<<'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
