# -*- coding: utf-8 -*-
r"""Is a double rule's white gap painted, or merely left alone?

Two readings hang on this. `28C006_3` puts a `right double` on a merged
block's far member and a `left thin` on the plain cell beside it, and Excel
draws the double CLEAN — no thin in its gap (SX95, where "a hollow rule wins
whichever side it is on" was tried and cost the `1c*zbd` five 0.054 each).
And `1c202304zbd` draws no vertical rule at the row where a horizontal double
sits, though the vertical rules above and below it are there (SX96).

Both follow if the double PAINTS its middle pixel instead of skipping it —
whatever was drawn there first is erased. So:

* a double across a FILLED cell: is the gap the fill's colour or white?
* a vertical rule crossing a horizontal double: does the crossing survive?
* the same with the double stated on the cell above instead of below, in case
  the order of drawing is what decides it.

    python tools\metrics\_xlsx_double_gap.py
    python tools\metrics\_xlsx_double_gap.py --reuse
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
SCRATCH = Path(r"C:\tmp\xlsx_double_gap")
FILL = 0x00FFFF          # BGR: a yellow nothing else in the sheet wears
ROW_PT = 24.0
COLUMN = 8.0
# (name, where the double is stated, whether the cells are filled)
ARMS = [
    ("double below, plain", "below", False),
    ("double below, filled", "below", True),
    ("double above, plain", "above", False),
    ("double above, filled", "above", True),
]


def build() -> Path | None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:H40").Interior.Color = 0xFFFFFF
        for column in (2, 3, 4):
            sheet.Columns(column).ColumnWidth = COLUMN
        at = 2
        for _name, where, filled in ARMS:
            for row in (at, at + 1):
                sheet.Rows(row).RowHeight = ROW_PT
                for column in (2, 3, 4):
                    cell = sheet.Cells(row, column)
                    if filled:
                        cell.Interior.Color = FILL
                    # Vertical rules through the whole pair, so one of them
                    # has to cross the double.
                    cell.Borders(7).LineStyle = 1     # xlEdgeLeft, continuous
                    cell.Borders(7).Weight = 2        # thin
                    cell.Borders(10).LineStyle = 1    # xlEdgeRight
                    cell.Borders(10).Weight = 2
            # The double on the boundary between the pair, stated by whichever
            # cell the arm names.
            block = (sheet.Range(sheet.Cells(at, 2), sheet.Cells(at, 4)) if where == "below"
                     else sheet.Range(sheet.Cells(at + 1, 2), sheet.Cells(at + 1, 4)))
            edge = 9 if where == "below" else 8       # xlEdgeBottom / xlEdgeTop
            block.Borders(edge).LineStyle = -4119     # xlDouble
            at += 3
        made = SCRATCH / "doublegap.xlsx"
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(2, 2), sheet.Cells(at, 4)).CopyPicture(
                    Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.8)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return made
        return None
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    if not args.reuse and build() is None:
        print("  Excel would not hand over a picture")
        return 1
    shot = np.asarray(Image.open(SCRATCH / "excel.png").convert("RGB"))
    tall = round(ROW_PT * 96 / 72)
    wide = round(COLUMN * 8 + 5)        # near enough; the columns are found by ink
    print(f"  picture {shot.shape[1]}x{shot.shape[0]}, rows of {tall}px")
    # The doubles are the rows with the most dark ink; read each one's three
    # rows and say what the middle holds where a vertical crosses it.
    dark = (shot.sum(axis=2) < 400)
    heavy = [y for y in range(shot.shape[0]) if dark[y].mean() > 0.5]
    print(f"  full-width dark rows: {heavy}")
    verticals = [x for x in range(shot.shape[1]) if dark[:, x].mean() > 0.5]
    print(f"  full-height dark columns: {verticals}")
    for at, (name, _where, _filled) in enumerate(ARMS):
        pair = [y for y in heavy if at * 3 * tall <= y < (at * 3 + 2) * tall + 4]
        if len(pair) < 2:
            print(f"  {name:<22} no double found ({pair})")
            continue
        gap = (pair[0] + pair[-1]) // 2
        held = [tuple(int(v) for v in shot[gap, x]) for x in verticals[:3]]
        away = tuple(int(v) for v in shot[gap, (verticals[0] + verticals[1]) // 2]) \
            if len(verticals) > 1 else None
        print(f"  {name:<22} double at {pair[0]}/{pair[-1]}, gap row {gap}:"
              f"  where the verticals cross {held}   between them {away}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
