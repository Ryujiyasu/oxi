"""How tall does Excel think a stacked block is, and where does it centre it?

`77df163c8f46_data_B01` fills a 117.9pt row with cells that carry
`textRotation="255"` (letters upright, stacked downward) and
`vertical="center"`. Excel puts the first letter's ink 19px below the row's
top; Oxi puts it 18. Every letter below inherits the pixel, and the whole
`data_A*`/`data_B*` family is built of these rows.

The pitch between letters already agrees, so the miss is in the block's own
height, or in how the half of the leftover is rounded. This asks Excel for
both at once: hold the letters, grow the row, and read where the first ink
lands. `block = row - 2 * offset` if the leftover is halved exactly, so a
block that comes out the same at every row height is a block Excel measured
once; a block that wobbles by a pixel with the row's parity is a rounding
rule instead.

Run: python tools/metrics/_xlsx_stack_center.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_stack_center")
ARMS = [("ＭＳ 明朝", 8.0), ("ＭＳ ゴシック", 8.0), ("ＭＳ Ｐゴシック", 11.0)]
WORDS = ["続柄", "性別番", "就業期間中", "職業分類番号を"]
HIGHS = [60.0, 75.0, 90.0, 105.0, 117.9]


def picture(sheet):
    for _ in range(8):
        try:
            sheet.Activate()
            sheet.Range("A1:F8").CopyPicture(Appearance=1, Format=2)
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
        sheet.Range("A1:F8").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = 4.0
        cell = sheet.Range("B3")
        cell.Orientation = -4166   # xlVertical: letters upright, stacked
        cell.VerticalAlignment = -4108     # xlCenter
        cell.WrapText = True
        print("  face          size  letters  row px   ink top-offset   implied block")
        for face, size in ARMS:
            cell.Font.Name = face
            cell.Font.Size = size
            for words in WORDS:
                cell.Value = words
                for high in HIGHS:
                    sheet.Rows("3").RowHeight = high
                    held = picture(sheet)
                    if held is None:
                        continue
                    top = round(cell.Top * 96 / 72)
                    left = round(cell.Left * 96 / 72)
                    high_px = round(cell.Height * 96 / 72)
                    wide_px = round(cell.Width * 96 / 72)
                    grey = np.asarray(held.convert("L"))[
                        top:top + high_px, left:left + wide_px]
                    ink = grey < 140
                    rows_lit = np.where(ink.any(axis=1))[0]
                    if not len(rows_lit):
                        print(f"  {face:<13}{size:>4.0f}  {words:<8} {high_px:>5}   no ink")
                        continue
                    first, last = int(rows_lit.min()), int(rows_lit.max())
                    block = high_px - 2 * first
                    print(f"  {face:<13}{size:>4.0f}  {words:<8} {high_px:>5}"
                          f"        +{first:<3} (ink {last - first + 1:>3} tall)"
                          f"      {block:>4}")
                    held.save(SCRATCH / f"{face}_{size:.0f}_{len(words)}_{high_px}.png")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
