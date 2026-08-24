# -*- coding: utf-8 -*-
"""Which way does Excel round the leftover when it centres text ACROSS a cell?

`_xlsx_center_round.py` asked this of the vertical leftover. The horizontal one
is what puts the floor workbook's right-hand column a pixel out:
`procurement-plan_outline_01` centres 「一般競争入札(総合評価)」 in a 296px
column and Excel starts it at 1004 where we start it at 1005 — same ink, same
width, one pixel of leftover rounded the other way.

The column is widened a pixel at a time, so the leftover walks through odd and
even and WHERE the offset steps is the rounding. Only column B is in the
picture, so the reading is the offset from the column's own left edge.

    python tools\\metrics\\_xlsx_center_across.py
"""

from __future__ import annotations

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
SCRATCH = Path(r"C:\tmp\xlsx_center_across")
# The floor workbook's own cell: full-width brackets, ＭＳ Ｐゴシック 12pt, in a
# column wide enough to hold it (a line that spills is centred across its
# neighbours instead, which is another rule).
WORDS = "契約の方法"
FACE, SIZE = "ＭＳ 明朝", 11.0
# A digit of the standard font is 7 pixels, so a fourteenth of a character is
# half a pixel: this walks the column's pixel width one at a time.
# Around the floor workbook's own 36.38 characters (296 pixels).
WIDTHS = [12.0 + step / 8.0 for step in range(14)]
ROW_PT = 18.0


def build(made: Path) -> int | None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:D{len(WIDTHS) + 4}").Interior.Color = 0xFFFFFF
        # One column can only hold one width, so every arm is its own SHEET
        # column and the picture is taken column by column. Simpler: one arm a
        # row is impossible here, so the width sweep runs across columns and
        # each is captured on its own.
        for at, width in enumerate(WIDTHS, start=2):
            sheet.Columns(at).ColumnWidth = width
            cell = sheet.Cells(2, at)
            cell.Value = WORDS
            cell.Font.Name = FACE
            cell.Font.Size = SIZE
            cell.HorizontalAlignment = -4108   # xlCenter
            cell.VerticalAlignment = -4108
            # The floor workbook's cell is ruled on all four sides, and a rule
            # may be what moves the leftover: half the arms are given one.
            cell.Borders.LineStyle = 1          # xlContinuous
            cell.Borders.Weight = 2             # xlThin
            # The floor workbook's cell WRAPS, and a wrapping cell is centred
            # against a different width: half the arms wrap.
            cell.WrapText = True
        sheet.Rows(2).RowHeight = ROW_PT
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(2, 2), sheet.Cells(2, 1 + len(WIDTHS))).CopyPicture(
                    Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.6)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                break
        else:
            return None
        return round(ROW_PT * 96 / 72)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def edges(made: Path) -> tuple[list[tuple[int, int, int]], int]:
    """Our own column spans, and where our row 2 starts."""
    import os
    told = subprocess.run([str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
                          env={"OXI_XLSX_DUMP_COLUMNS": "1", "OXI_XLSX_DUMP_ROWS": "1",
                               **os.environ},
                          capture_output=True, text=True, encoding="utf-8")
    out, at, down, top = [], 0, 0, 0
    for line in (told.stdout or "").splitlines() + (told.stderr or "").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "column":
            out.append((int(parts[1]), at, at + int(parts[3])))
            at += int(parts[3])
        if len(parts) == 4 and parts[0] == "row":
            if int(parts[1]) == 2:
                top = down
            down += int(parts[3])
    return out, top


def main() -> int:
    made = SCRATCH / "centre.xlsx"
    tall = build(made)
    if tall is None:
        print("  Excel would not hand over a picture")
        return 1
    columns, ours_top = edges(made)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    ours = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140
    # The picture starts at sheet column B, which the dump numbers 1.
    spans = {index: (left, right) for index, left, right in columns}
    print(f"  {FACE} {SIZE}pt, {WORDS!r} centred")
    print("  column px   ink-left: Excel Oxi   leftover   where the odd pixel goes")
    walked = 0
    for at in range(len(WIDTHS)):
        index = 1 + at                       # sheet column B, C, D, ...
        if index not in spans:
            continue
        left, right = spans[index]
        wide = right - left
        # A ruled cell's own border is ink at both edges, so the reading skips
        # two pixels either side and the top and foot of the band.
        band = truth[3:tall - 3, walked + 2:walked + wide - 2]
        lit = np.where(band.any(axis=0))[0]
        theirs = int(lit.min()) + 2 if len(lit) else None
        band = ours[ours_top + 3:ours_top + tall - 3, left + 2:right - 2]
        lit = np.where(band.any(axis=0))[0]
        mine = int(lit.min()) + 2 if len(lit) else None
        walked += wide
        if theirs is None or mine is None:
            print(f"  {wide:>9}   nothing to read")
            continue
        print(f"  {wide:>9} {'wrap':<8}"
              f"   {theirs:>5} {mine:>5}"
              f"      {'odd' if wide % 2 else 'even':>4}"
              f"       {'ours right' if mine > theirs else 'ours left' if mine < theirs else 'same'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
