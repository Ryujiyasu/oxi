# -*- coding: utf-8 -*-
r"""Which way does a WRAPPED cell's block round when the row centres it?

`_xlsx_center_round.py` put this question for a single line and read a floor,
over thirteen row heights and eight faces, and that is what the renderer does.
The `h2daa*_dendeba_kmc` trio says the answer is different once the cell wraps:
its `C29` is `vertical="center" wrapText="1"`, two lines of ＭＳ Ｐゴシック 11pt
in a 33px row, and every line of it — and the rule under it — sits ONE pixel
high in our picture. Both row tops match Excel's to the pixel, so what is out
is the halving of the leftover inside the row.

That is the same split SX85 found across the cell: a wrapping cell rounds the
other way from one that does not. So the sweep is by line count — one, two and
three lines in the same column — with the row walked a pixel at a time so the
leftover runs through odd and even.

    python tools\metrics\_xlsx_wrap_center_round.py
    python tools\metrics\_xlsx_wrap_center_round.py --reuse
"""

from __future__ import annotations

import argparse
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
SCRATCH = Path(r"C:\tmp\xlsx_wrap_center_tight"
               if "--tight" in sys.argv
               else r"C:\tmp\xlsx_wrap_center")
FACE, POINTS = "ＭＳ Ｐゴシック", 11.0
COLUMN = 9.0
# One, two and three lines in a column that holds about ten half-width letters.
WORDS = {1: "Electronic", 2: "Electronic tubes", 3: "Electronic tubes devices and"}
# Row heights in points chosen so the pixel height walks one at a time, wide
# enough to hold three lines and then some.
HIGHS = [round(px * 72 / 96, 2) for px in (range(10, 30) if "--tight" in sys.argv
                                          else range(30, 52))]


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:F{len(HIGHS) + 4}").Interior.Color = 0xFFFFFF
        for column in (2, 3, 4):
            sheet.Columns(column).ColumnWidth = COLUMN
        for at, high in enumerate(HIGHS, start=2):
            for column, count in ((2, 1), (3, 2), (4, 3)):
                cell = sheet.Cells(at, column)
                cell.Value = WORDS[count]
                cell.Font.Name = FACE
                cell.Font.Size = POINTS
                cell.WrapText = True
                cell.HorizontalAlignment = -4131      # left
                cell.VerticalAlignment = -4108        # centre
            sheet.Rows(at).RowHeight = high
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(2, 2), sheet.Cells(1 + len(HIGHS), 4)).CopyPicture(
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


def ours(made: Path):
    told = subprocess.run([str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
                          env={"OXI_XLSX_DUMP_COLUMNS": "1", "OXI_XLSX_DUMP_ROWS": "1",
                               **os.environ},
                          capture_output=True, text=True, encoding="utf-8")
    columns, rows, at, down = {}, {}, 0, 0
    for line in (told.stdout or "").splitlines() + (told.stderr or "").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "column":
            columns[int(parts[1])] = (at, at + int(parts[3]))
            at += int(parts[3])
        if len(parts) == 4 and parts[0] == "row":
            rows[int(parts[1])] = (down, down + int(parts[3]))
            down += int(parts[3])
    return np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140, columns, rows


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    parser.add_argument("--tight", action="store_true",
                        help="rows too short to hold even one line")
    args = parser.parse_args()
    made = SCRATCH / "wrapcentre.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    mine, columns, rows = ours(made)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    wide = columns[1][1] - columns[1][0]
    print(f"  {FACE} {POINTS}pt, wrapped and centred in a {wide}px column")
    print("  row px | 1 line  Excel   Oxi   | 2 lines Excel   Oxi   | 3 lines Excel   Oxi")
    walked = 0
    for at, _high in enumerate(HIGHS):
        top, foot = rows[at + 2]
        tall = foot - top
        told = []
        for column in (0, 1, 2):
            # The reading skips a pixel at each edge, where the sheet's own
            # gridline is ink of its own.
            band = truth[walked + 1:walked + tall - 1,
                         column * wide + 1:(column + 1) * wide - 1]
            lit = np.where(band.any(axis=1))[0]
            theirs = (int(lit.min()) + 1, int(lit.max()) + 1) if len(lit) else None
            left, right = columns[column + 1]
            band = mine[top + 1:foot - 1, left + 1:right - 1]
            lit = np.where(band.any(axis=1))[0]
            ours_at = (int(lit.min()) + 1, int(lit.max()) + 1) if len(lit) else None
            told.append((theirs, ours_at))
        walked += tall
        if any(one is None for pair in told for one in pair):
            print(f"  {tall:>6} |  nothing to read")
            continue
        print(f"  {tall:>6} |"
              + "".join(f" {a[0]:>3}-{a[1]:<3} {b[0]:>3}-{b[1]:<3}{'' if a == b else '<<':<2}|"
                        for a, b in told))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
