# -*- coding: utf-8 -*-
r"""Where does a shape land when its anchor is deep into a sheet of narrow
columns?

`_xlsx_shape_origin.py` sweeps a shape's left in column A and our box lands on
Excel's in all sixteen arms. `002` still puts its memo shapes a pixel out —
sometimes left, sometimes right — and those are anchored at columns 57 and 59
of a sheet whose columns are 1.5 and 2.4 characters wide, so the edge the
anchor counts from is a SUM of fractional widths. Rounding each column and
adding is not the same number as adding and rounding once, and the difference
only shows once enough of them are behind you.

So: columns of a stated fractional width, a shape anchored at the far end of a
run of them, and the run's length swept. The reading is the shape's own left
and right edge, ours beside Excel's; the shape is filled and unlettered so its
edges are the ink.

    python tools\metrics\_xlsx_shape_far_anchor.py
    python tools\metrics\_xlsx_shape_far_anchor.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_far_anchor")
# `002`'s own column widths, and one that lands on a half pixel.
COLUMN = 2.42578125
# The anchor column of each arm: far enough apart that the sums differ by a
# fraction of a pixel each time.
ANCHORS = list(range(2, 60, 4))
WIDE, HIGH = 60.0, 14.0
ROW_PT = 18.0


def build(made: Path) -> list[tuple[int, float, float]]:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    placed = []
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:BZ60").Interior.Color = 0xFFFFFF
        sheet.Columns("A:BZ").ColumnWidth = COLUMN
        at = 2
        for column in ANCHORS:
            sheet.Rows(at).RowHeight = ROW_PT
            sheet.Rows(at + 1).RowHeight = ROW_PT
            # AddShape wants points; the cell's own Left is the anchor Excel
            # will write, so the shape is put exactly on it.
            cell = sheet.Cells(at, column)
            box = sheet.Shapes.AddShape(1, cell.Left, cell.Top, WIDE, HIGH)
            box.Fill.Visible = True
            box.Fill.ForeColor.RGB = 0
            box.Line.Visible = False
            try:
                box.Shadow.Visible = False
            except Exception:
                pass
            box.TextFrame2.TextRange.Text = ""
            placed.append((column, box.Left, box.Top))
            at += 2
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range("A1:BZ60").CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.9)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return placed
        return []
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def edges(picture: np.ndarray, top: float) -> tuple[int, int] | None:
    """The filled shape's own left and right, in its own band."""
    start = max(0, round(top * 96 / 72) + 2)
    band = picture[start:start + round(HIGH * 96 / 72) - 4]
    lit = np.where((band < 120).any(axis=0))[0]
    if not len(lit):
        return None
    return int(lit.min()), int(lit.max())


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "deep.xlsx"
    if args.reuse:
        placed = []
        at = 2
        for column in ANCHORS:
            placed.append((column, 0.0, (at - 1) * ROW_PT))
            at += 2
    else:
        placed = build(made)
        if not placed:
            print("  Excel would not hand over a picture")
            return 1
    subprocess.run([str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
                   env={**os.environ}, capture_output=True, check=False)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L"))
    mine = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L"))
    print(f"  columns of {COLUMN} characters, a {WIDE}pt shape on each anchor")
    print("  anchor  left pt   |  box: Excel        Oxi")
    for column, left, top in placed:
        theirs = edges(truth, top)
        ours = edges(mine, top)
        if theirs is None or ours is None:
            print(f"  {column:>6}   nothing to read ({theirs} / {ours})")
            continue
        print(f"  {column:>6} {left:>9.3f}   {theirs[0]:>5}-{theirs[1]:<5}"
              f" {ours[0]:>7}-{ours[1]:<5}"
              f"  {'' if theirs == ours else '<<'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
