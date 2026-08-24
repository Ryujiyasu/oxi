"""Does Excel centre a stacked cell's LINE BOXES, or its INK?

`_xlsx_stack_block.py` says the block Excel centres is `(n-1) * pitch + one`,
and that `one` is not the pitch: 10 against a pitch of 14 at 8pt, 12 against
18 at 11pt. Every one of those `one` values came out exactly one more than the
ink of the single letter that was used, which would mean Excel centres what
the letters actually cover rather than the boxes they sit in.

If that is so, a letter with almost no ink must move the whole stack: 「一」 is
one stroke, 「■」 fills its em. If instead the boxes are centred, the letter
cannot matter at all. Same face, same size, same count, same row — only the
letter changes.

Run: python tools/metrics/_xlsx_stack_ink.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_stack_ink")
FACE = "ＭＳ ゴシック"
SIZES = [8.0, 11.0]
LETTERS = ["一", "続", "■", "あ", "ー"]
COUNTS = [1, 2]
ROWS_PX = [120, 121]


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:D120").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = 5.0
        plan, at = [], 2
        for size in SIZES:
            for letter in LETTERS:
                for count in COUNTS:
                    for row_px in ROWS_PX:
                        cell = sheet.Cells(at, 2)
                        cell.Value = letter * count
                        cell.Font.Name = FACE
                        cell.Font.Size = size
                        cell.Orientation = -4166
                        cell.VerticalAlignment = -4108
                        cell.WrapText = True
                        sheet.Rows(at).RowHeight = round(row_px * 72 / 96, 2)
                        plan.append((at, size, letter, count, row_px))
                        at += 1
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(at - 1, 2))
        for _ in range(10):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(0.8)
            shot = ImageGrab.grabclipboard()
            if shot is not None:
                break
        else:
            print("Excel would not hand over a picture")
            return 1
        shot.save(SCRATCH / "shot.png")
        grey = np.asarray(shot.convert("L"))
        base = round(sheet.Cells(2, 2).Top * 96 / 72)
        wide = round(sheet.Cells(2, 2).Width * 96 / 72)
        seen = {}
        for row, size, letter, count, row_px in plan:
            top = round(sheet.Cells(row, 2).Top * 96 / 72) - base
            high = round(sheet.Cells(row, 2).Height * 96 / 72)
            block = grey[top:top + high, 0:wide] < 140
            lit = np.where(block.any(axis=1))[0]
            seen[(size, letter, count, row_px)] = (
                (int(lit.min()), int(lit.max() - lit.min() + 1)) if len(lit) else None)
        print(f"  {FACE}")
        print("  size  letter  n   offset   ink   implied block   block - ink")
        for size in SIZES:
            for letter in LETTERS:
                for count in COUNTS:
                    at = seen.get((size, letter, count, 120))
                    nxt = seen.get((size, letter, count, 121))
                    if at is None or nxt is None:
                        print(f"  {size:>4.0f}  {letter:<6}  {count}   no ink")
                        continue
                    off, ink = at
                    block = (121 - 2 * off) if nxt[0] == off else (120 - 2 * off)
                    print(f"  {size:>4.0f}  {letter:<6}  {count}   {off:>6}"
                          f"   {ink:>3}   {block:>13}   {block - ink:>11}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
