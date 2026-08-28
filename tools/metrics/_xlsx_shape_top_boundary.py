# -*- coding: utf-8 -*-
r"""Exactly where does a shape's first line step down a pixel?

`_xlsx_shape_origin_down.py` says Excel adds the exact inset to the exact top
edge and rounds once — 8 arms of 8, five shapes each. `tb_r8_jizensoudan`
says otherwise: its panel's top is 8.6865 pixels, its inset 4.8, and the sum
13.4865 should round to 13, where Excel draws it at 14. Fourteen thousandths
of a pixel decide it, which is not a difference an argument can settle.

So: put the boundary itself under the sweep. The top edge is stepped a
hundredth of a pixel at a time across the place where the sum crosses a half,
with a reference box at a whole pixel in every band to read against, and the
same sweep run for a bordered shape as well as a bare one.

  * a step down inside the sweep says the exact top is what Excel rounds, and
    WHERE it steps says whether the half is the boundary;
  * no step at all says the top was already on a whole pixel before the inset
    was added to it.

    python tools\metrics\_xlsx_shape_top_boundary.py
    python tools\metrics\_xlsx_shape_top_boundary.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_top_boundary")
LETTER = "日"
# The face, size and border the corpus panel that disagrees is set in.
FACE, POINTS = "游ゴシック", 12.0
# A band 45 points tall is 60 pixels, so every band begins on a whole pixel
# and the reference box in it does too.
BAND = 45.0
WIDE, TALL = 120.0, 30.0
# The first sweep found Excel stepping down at a sum-fraction of 0.47, not
# 0.50 — so something adds about three hundredths of a pixel. Whether that
# belongs to the INSET or to the rounding of the top is what three lanes
# answer: their insets differ by 0.8 and 0.6 of a pixel, so if all three step
# at the same SUM fraction the bias is in the rounding, and if they step at
# the same TOP fraction it is not the sum being rounded at all.
MARGINS = [0.0, 3.6, 7.2]       # points; 0, 4.8 and 9.6 pixels
# The three lanes stepped at 0.47, 0.67 and 0.87 — the same SUM fraction each
# time, so the sum is what Excel rounds and its boundary is not the half. This
# narrows that boundary to a thousandth: 0.46875 would be a thirty-second of a
# pixel of bias, 0.4667 a fifteenth of a tenth, and the two are worth telling
# apart. Only the first lane's boundary is inside this sweep; the other two
# hold still, which is the check that nothing else moved.
OFFSETS = [0.455 + step * 0.001 for step in range(26)]


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:P400").Interior.Color = 0xFFFFFF
        for at, offset in enumerate(OFFSETS):
            here = 15.0 + at * BAND
            # The reference: a filled box on a whole pixel, so what the text
            # is read against does not move with the arm.
            mark = sheet.Shapes.AddShape(1, 20.0, here, 40.0, 12.0)
            mark.Fill.Visible = True
            mark.Fill.ForeColor.RGB = 0
            mark.Line.Visible = False
            for lane, margin in enumerate(MARGINS):
                words = sheet.Shapes.AddShape(
                    1, 120.0 + lane * 200.0, here, WIDE, TALL)
                words.Fill.Visible = False
                words.Line.Visible = False
                frame = words.TextFrame2
                frame.MarginTop = margin
                frame.WordWrap = False
                frame.AutoSize = 0
                frame.VerticalAnchor = 1          # top
                frame.TextRange.Text = LETTER
                frame.TextRange.Font.Size = POINTS
                frame.TextRange.Font.Name = FACE
                frame.TextRange.Font.NameFarEast = FACE
                # A shape with no fill takes its theme's text colour, which
                # is white: say black or the reader finds nothing.
                frame.TextRange.Font.Fill.ForeColor.RGB = 0
                # The offset is a fraction of a PIXEL, and a pixel is 0.75pt.
                words.Top = here + offset * 0.75
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(1, 1), sheet.Cells(
                    int(len(OFFSETS) * BAND / 15) + 12, 30)).CopyPicture(
                    Appearance=1, Format=2)
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


def first_ink(dark: np.ndarray, y0: int, y1: int, x0: int, x1: int) -> int | None:
    for y in range(y0, min(y1, dark.shape[0])):
        if dark[y, x0:x1].sum() >= 2:
            return y
    return None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "boundary.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    subprocess.run(
        [str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
        capture_output=True, text=True, encoding="utf-8", env=dict(os.environ),
    )
    mine = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140
    band = round(BAND * 96 / 72)
    print(f"  {'top +px':>9}   {'Excel mark/0/3.6/7.2':>26}"
          f"{'Oxi mark/0/3.6/7.2':>26}")
    for at, offset in enumerate(OFFSETS):
        y0, y1 = at * band, (at + 1) * band
        held = []
        for dark in (truth, mine):
            mark = first_ink(dark, y0, y1, 30, 70)
            lanes = [first_ink(dark, y0, y1, 165 + lane * 267, 290 + lane * 267)
                     for lane in range(len(MARGINS))]
            held.append((mark, *[None if mark is None or one is None else one - mark
                                 for one in lanes]))
        same = held[0] == held[1]
        print(f"  {offset:>9.2f}   {str(held[0]):>26}"
              f"{str(held[1]):>26}{'' if same else '  <<'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
