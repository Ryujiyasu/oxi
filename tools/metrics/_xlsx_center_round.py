"""Which way does Excel round the leftover when it centres a cell's text?

The stacked sweep says the spare pixel goes ABOVE: a 157px row holding a 24px
block puts the first ink at 67, not 66. Oxi floors instead, and its stacked
block is two pixels short, so the two mistakes cancel on odd rows and show on
even ones. Before the shared centring is touched, the same question has to be
put for ORDINARY text, whose floor was measured over thirteen row heights and
eight fonts and is what the current code does.

Row heights are swept one pixel at a time. A block of fixed height centred in
a growing row steps its offset by one every two pixels; WHERE the step lands
is the rounding. Both kinds are measured in the same sheet so the two answers
are comparable.

Run: python tools/metrics/_xlsx_center_round.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_center_round")
FACE, SIZE = "ＭＳ 明朝", 8.0
# Row heights in points chosen so the pixel height walks 100..111 one at a time.
HIGHS = [round(px * 72 / 96, 2) for px in range(100, 112)]


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:F40").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = 4.0
        sheet.Columns("D").ColumnWidth = 8.0
        rows = []
        for at, high in enumerate(HIGHS, start=2):
            stacked = sheet.Cells(at, 2)
            stacked.Value = "続柄"
            stacked.Font.Name = FACE
            stacked.Font.Size = SIZE
            stacked.Orientation = -4166
            stacked.VerticalAlignment = -4108
            stacked.WrapText = True
            plain = sheet.Cells(at, 4)
            plain.Value = "あ"
            plain.Font.Name = FACE
            plain.Font.Size = SIZE
            plain.VerticalAlignment = -4108
            sheet.Rows(at).RowHeight = high
            rows.append(at)
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(rows[-1], 4))
        for _ in range(8):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.5)
            shot = ImageGrab.grabclipboard()
            if shot is not None:
                break
        else:
            print("Excel would not hand over a picture")
            return 1
        shot.save(SCRATCH / "shot.png")
        grey = np.asarray(shot.convert("L"))
        base = round(sheet.Cells(2, 2).Top * 96 / 72)
        stack_left = 0
        stack_wide = round(sheet.Cells(2, 2).Width * 96 / 72)
        plain_left = round(sheet.Cells(2, 4).Left * 96 / 72) - round(
            sheet.Cells(2, 2).Left * 96 / 72)
        plain_wide = round(sheet.Cells(2, 4).Width * 96 / 72)
        print(f"  {FACE} {SIZE:.0f}pt")
        print("  row px    stacked offset / ink   plain offset / ink")
        for at in rows:
            top = round(sheet.Cells(at, 2).Top * 96 / 72) - base
            high = round(sheet.Cells(at, 2).Height * 96 / 72)
            out = []
            for left, wide in ((stack_left, stack_wide), (plain_left, plain_wide)):
                blk = grey[top:top + high, left:left + wide] < 140
                lit = np.where(blk.any(axis=1))[0]
                out.append((int(lit.min()), int(lit.max() - lit.min() + 1))
                           if len(lit) else (None, None))
            (so, si), (po, pi) = out
            print(f"  {high:>6}      {str(so):>5} / {str(si):<4}"
                  f"        {str(po):>5} / {str(pi):<4}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
