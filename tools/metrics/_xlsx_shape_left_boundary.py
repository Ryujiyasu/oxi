# -*- coding: utf-8 -*-
r"""Does the side of a shape's text turn on the same sixteenth the top does?

`_xlsx_shape_top_boundary.py` found the top edge and its inset added together
and put on a pixel at a boundary of 15/32, not a half — Excel snaps the sum to
a SIXTEENTH of a pixel first. The sides were settled before that, by
`_xlsx_shape_origin.py`, at quarter-pixel steps: a step that coarse cannot see
a thirty-second of bias either way, so the question is open and the same sweep
answers it.

Two lanes an arm, one above the other so both share the swept Left: the inset
written 0 and 7.2 points. If both step at the same SUM fraction the sum is
what is rounded, and where they step says whether the boundary is the half or
the sixteenth.

    python tools\metrics\_xlsx_shape_left_boundary.py
    python tools\metrics\_xlsx_shape_left_boundary.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_left_boundary")
LETTER = "日"
FACE, POINTS = "游ゴシック", 12.0
# A lane 30 points tall is 40 pixels and an arm is two of them, so every arm
# begins on a whole pixel and the reference box in it does too.
LANE = 37.5
ARM = 75.0
WIDE, TALL = 120.0, 28.0
LEFT = 100.0                    # where the text shapes stand, in points
MARGINS = [0.0, 7.2]            # points; 0 and 9.6 pixels
# The coarse sweep stepped at 0.14 and 0.54 — 0.40 apart, which is exactly
# what a sum being rounded predicts when the second lane's inset adds 0.6 of a
# pixel. The shapes stand at 100 points = 133.333 pixels, so the first lane's
# step is at a sum fraction of 0.473. This narrows it: a sixteenth's boundary
# is 0.46875 (offset 0.1354) and a twip's is 0.46667 (offset 0.1333).
OFFSETS = [0.130 + step * 0.001 for step in range(12)]


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:P800").Interior.Color = 0xFFFFFF
        for at, offset in enumerate(OFFSETS):
            here = 15.0 + at * ARM
            # The reference: a filled box on a whole pixel, tall enough to
            # stand beside both lanes, so what the text is read against does
            # not move with the arm.
            mark = sheet.Shapes.AddShape(1, 20.0, here, 12.0, ARM - 6.0)
            mark.Fill.Visible = True
            mark.Fill.ForeColor.RGB = 0
            mark.Line.Visible = False
            for lane, margin in enumerate(MARGINS):
                words = sheet.Shapes.AddShape(
                    1, LEFT, here + lane * LANE, WIDE, TALL)
                words.Fill.Visible = False
                words.Line.Visible = False
                frame = words.TextFrame2
                frame.MarginLeft = margin
                frame.WordWrap = False
                frame.AutoSize = 0
                frame.HorizontalAnchor = 1        # left
                frame.VerticalAnchor = 1          # top
                frame.TextRange.ParagraphFormat.Alignment = 1
                frame.TextRange.Text = LETTER
                frame.TextRange.Font.Size = POINTS
                frame.TextRange.Font.Name = FACE
                frame.TextRange.Font.NameFarEast = FACE
                # A shape with no fill takes its theme's text colour, which is
                # white: say black or the reader finds nothing.
                frame.TextRange.Font.Fill.ForeColor.RGB = 0
                # The offset is a fraction of a PIXEL, and a pixel is 0.75pt.
                words.Left = LEFT + offset * 0.75
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(1, 1), sheet.Cells(
                    int(len(OFFSETS) * ARM / 15) + 12, 30)).CopyPicture(
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
    """The first COLUMN holding ink in a window."""
    held = dark[y0:min(y1, dark.shape[0]), x0:x1]
    for x in range(held.shape[1]):
        if held[:, x].sum() >= 2:
            return x0 + x
    return None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "leftboundary.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    subprocess.run(
        [str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
        capture_output=True, text=True, encoding="utf-8", env=dict(os.environ),
    )
    mine = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140
    arm = round(ARM * 96 / 72)
    lane = round(LANE * 96 / 72)
    print(f"  {'left +px':>10}   {'Excel mark/0/7.2':>22}{'Oxi mark/0/7.2':>22}")
    for at, offset in enumerate(OFFSETS):
        y0 = at * arm
        held = []
        for dark in (truth, mine):
            mark = first_ink(dark, y0, y0 + arm, 20, 60)
            # Each lane is read in a window pinned to its OWN shape, not to an
            # even share of the arm: a line of 12 point text is 27 pixels tall
            # and the first attempt let the lane above bleed into the one
            # below, which read as a column of ink a pixel to its left.
            lanes = [first_ink(dark, y0 + 23 + step * lane, y0 + 47 + step * lane,
                               120, 400)
                     for step in range(len(MARGINS))]
            held.append((mark, *[None if mark is None or one is None else one - mark
                                 for one in lanes]))
        same = held[0] == held[1]
        print(f"  {offset:>10.2f}   {str(held[0]):>22}{str(held[1]):>22}"
              f"{'' if same else '  <<'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
