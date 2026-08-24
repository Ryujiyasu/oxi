"""How tall is a stacked block, per face and per size?

A stacked cell is centred by `offset = ceil((row - block) / 2)`, which is the
same thing Oxi's `floor((row - b) / 2)` does when `b = block - 1`. So the only
thing that has to be right is the block, and Oxi's is one pixel short at 8pt
and one pixel long at 11pt — a constant nudge would fix one family of books
and break the other.

The block is read off Excel two row heights at a time. For an even leftover
the halving is exact; for an odd one the spare pixel goes above. Measuring at
`row` and `row + 1` tells the parity apart:

    offset(row + 1) == offset(row)  ->  block is odd,  block = row + 1 - 2*offset
    otherwise                       ->  block is even, block = row - 2*offset

With the block known at two letter counts, `pitch = block(n+1) - block(n)` and
the height of one letter's box is `block(n) - (n - 1) * pitch`.

Run: python tools/metrics/_xlsx_stack_block.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_stack_block")
FACES = ["ＭＳ 明朝", "ＭＳ ゴシック", "ＭＳ Ｐゴシック", "ＭＳ Ｐ明朝"]
SIZES = [8.0, 9.0, 10.0, 10.5, 11.0, 12.0, 14.0]
COUNTS = [1, 2, 3, 4]
ROWS_PX = [120, 121, 122, 123]
LETTER = "続"


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:D200").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = 5.0
        plan, at = [], 2
        for face in FACES:
            for size in SIZES:
                for count in COUNTS:
                    for row_px in ROWS_PX:
                        cell = sheet.Cells(at, 2)
                        cell.Value = LETTER * count
                        cell.Font.Name = face
                        cell.Font.Size = size
                        cell.Orientation = -4166           # xlVertical
                        cell.VerticalAlignment = -4108     # xlCenter
                        cell.WrapText = True
                        sheet.Rows(at).RowHeight = round(row_px * 72 / 96, 2)
                        plan.append((at, face, size, count, row_px))
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
        seen: dict[tuple, int] = {}
        for row, face, size, count, row_px in plan:
            top = round(sheet.Cells(row, 2).Top * 96 / 72) - base
            high = round(sheet.Cells(row, 2).Height * 96 / 72)
            block = grey[top:top + high, 0:wide] < 140
            lit = np.where(block.any(axis=1))[0]
            seen[(face, size, count, row_px)] = int(lit.min()) if len(lit) else None
        print("  face          size   block for 1/2/3/4 letters   pitch   agrees")
        for face in FACES:
            for size in SIZES:
                found, steady = {}, True
                for count in COUNTS:
                    guesses = set()
                    for row_px in ROWS_PX:
                        at = seen.get((face, size, count, row_px))
                        nxt = seen.get((face, size, count, row_px + 1))
                        if at is None or nxt is None:
                            continue
                        guesses.add((row_px + 1 - 2 * at) if nxt == at else (row_px - 2 * at))
                    if len(guesses) == 1:
                        found[count] = guesses.pop()
                    else:
                        found[count] = sorted(guesses) if guesses else None
                        steady = False
                blocks = [found.get(c) for c in COUNTS]
                steps = [b - a for a, b in zip(blocks, blocks[1:])
                         if isinstance(a, int) and isinstance(b, int)]
                pitch = steps[0] if steps and len(set(steps)) == 1 else steps
                print(f"  {face:<13}{size:>5.1f}   {str(blocks):<26}"
                      f" {str(pitch):>7}   {'yes' if steady else 'NO'}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
