# -*- coding: utf-8 -*-
r"""Is a shape's box the size we think it is?

`zuhyo`'s three-line note only comes out right if its box is a pixel shorter
than we compute: with the exact pitch and a 62-pixel box the third line lands
at 568, with a 61-pixel box it lands at 567, which is Excel's. The rounding
shipped for it — starting the block on a whole pixel — reproduces the same
answer, so one of the two is a compensation for the other.

Rather than reason about which, this asks Excel for the box itself. A
rectangle an arm, anchored at a cell and given a size, drawn with a visible
outline: the answer is where its edges land, in Excel's picture and in ours.
A box that agrees leaves the pitch as the only suspect; a box that does not is
the whole story.

    python tools\metrics\_xlsx_shape_box.py
    python tools\metrics\_xlsx_shape_box.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_box")
# Sizes in points, chosen so the pixel answer lands on and off whole numbers:
# 46.5pt is 62px exactly, 46.125pt is 61.5, 45.75pt is 61.
SIZES = [30.0, 45.75, 46.125, 46.5, 60.0, 60.375, 75.0, 82.5]
WIDE = 90.0
STEP = 100.0


def build() -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:J300").Interior.Color = 0xFFFFFF
        for at, tall in enumerate(SIZES):
            box = sheet.Shapes.AddShape(1, 20.0, 10.0 + at * STEP, WIDE, tall)
            box.Fill.Visible = False
            box.Line.Visible = True
            box.Line.ForeColor.RGB = 0
            box.Line.Weight = 0.75
        book.SaveAs(str(SCRATCH / "boxes.xlsx"), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(1, 1), sheet.Cells(
                    int(len(SIZES) * STEP / 15) + 12, 10)).CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(1.0)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def rules(dark: np.ndarray, y0: int, y1: int) -> tuple[int, int] | None:
    """The top and bottom edges of the one rectangle in this band."""
    rows = [y for y in range(y0, min(y1, dark.shape[0])) if dark[y].sum() >= 40]
    return (rows[0], rows[-1]) if len(rows) >= 2 else None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    if not args.reuse and not build():
        print("  Excel would not hand over a picture")
        return 1
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    subprocess.run(
        [str(RENDERER), str(SCRATCH / "boxes.xlsx"), str(SCRATCH / "oxi.png"), "96"],
        capture_output=True, text=True, encoding="utf-8", env=dict(os.environ),
    )
    mine = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140
    band = round(STEP * 96 / 72)
    print(f"  {'pt':>9}{'96dpi px':>10}{'Excel top..foot':>18}{'tall':>6}"
          f"{'Oxi top..foot':>16}{'tall':>6}")
    agree = 0
    for at, tall in enumerate(SIZES):
        one = rules(truth, at * band, (at + 1) * band)
        two = rules(mine, at * band, (at + 1) * band)
        if one is None or two is None:
            print(f"  {tall:>9}{tall * 96 / 72:>10.2f}   nothing to read")
            continue
        same = (one[1] - one[0]) == (two[1] - two[0]) and one[0] == two[0]
        agree += same
        print(f"  {tall:>9}{tall * 96 / 72:>10.2f}{f'{one[0]}..{one[1]}':>18}"
              f"{one[1] - one[0] + 1:>6}{f'{two[0]}..{two[1]}':>16}"
              f"{two[1] - two[0] + 1:>6}{'' if same else '  <<'}")
    print(f"  {agree} of {len(SIZES)} boxes agree")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
