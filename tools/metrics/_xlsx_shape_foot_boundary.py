# -*- coding: utf-8 -*-
r"""Where does the FOOT of a shape's text box turn over to the next pixel?

SX119 measured the top: Excel adds the inset to the exact edge, snaps the sum
to a sixteenth of a pixel and rounds, which puts the boundary at 15/32 rather
than a half. SX123 then gave the foot the same treatment on the assumption
that it mirrors the top — an assumption, not a measurement, and the foot only
matters for a block anchored `ctr` or `b`, which is where `glossary_05`'s
panel disagrees. Solving that panel's twelve arms together says its area wants
to be a pixel taller than we compute.

So ask. A block anchored `b` hangs from the foot, so its ink follows the foot
one for one: sweep the box's HEIGHT a hundredth of a pixel at a time with the
top pinned, and the step in the ink is the step in the foot. Two lanes whose
bottom insets differ by 0.6 of a pixel say whether it is the sum being
rounded, exactly as they did for the top.

    python tools\metrics\_xlsx_shape_foot_boundary.py
    python tools\metrics\_xlsx_shape_foot_boundary.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_foot_boundary")
LETTER = "日"
FACE, POINTS = "游ゴシック", 12.0
# A band 45 points tall is 60 pixels, so every arm begins on a whole pixel and
# the reference box in it does too.
BAND = 45.0
WIDE = 120.0
TALL = 30.0             # the box's height before the sweep's fraction is added
MARGINS = [0.0, 7.2]    # bottom inset in points; 0 and 9.6 pixels
OFFSETS = [step * 0.01 for step in range(100)]


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:Z2000").Interior.Color = 0xFFFFFF
        for at, offset in enumerate(OFFSETS):
            here = 15.0 + at * BAND
            # The reference: a filled box on a whole pixel, so what the text is
            # read against does not move with the arm.
            mark = sheet.Shapes.AddShape(1, 20.0, here, 40.0, 12.0)
            mark.Fill.Visible = True
            mark.Fill.ForeColor.RGB = 0
            mark.Line.Visible = False
            mark.Top = here
            for lane, margin in enumerate(MARGINS):
                words = sheet.Shapes.AddShape(
                    1, 120.0 + lane * 200.0, here, WIDE, TALL)
                words.Fill.Visible = False
                words.Line.Visible = False
                frame = words.TextFrame2
                frame.MarginBottom = margin
                frame.MarginTop = 0.0
                frame.WordWrap = False
                frame.AutoSize = 0
                # Hung from the FOOT, which is the whole question.
                frame.VerticalAnchor = 4          # bottom
                frame.TextRange.Text = LETTER
                frame.TextRange.Font.Size = POINTS
                frame.TextRange.Font.Name = FACE
                frame.TextRange.Font.NameFarEast = FACE
                # A shape with no fill takes its theme's text colour, which is
                # white: say black or the reader finds nothing.
                frame.TextRange.Font.Fill.ForeColor.RGB = 0
                words.Top = here
                # The offset is a fraction of a PIXEL, and a pixel is 0.75pt.
                words.Height = TALL + offset * 0.75
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


def edge(dark: np.ndarray, y0: int, y1: int, x0: int, x1: int, last: bool) -> int | None:
    held = dark[y0:min(y1, dark.shape[0]), x0:x1]
    rows = np.where(held.sum(axis=1) >= 2)[0]
    if len(rows) == 0:
        return None
    return y0 + int(rows[-1] if last else rows[0])


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "footboundary.xlsx"
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
    print(f"  {'height +px':>11}   {'Excel mark/0/7.2':>22}{'Oxi mark/0/7.2':>22}")
    for at, offset in enumerate(OFFSETS):
        y0 = at * band
        held = []
        for dark in (truth, mine):
            mark = edge(dark, y0, y0 + band, 30, 70, False)
            lanes = [edge(dark, y0, y0 + band, 165 + lane * 267, 290 + lane * 267, True)
                     for lane in range(len(MARGINS))]
            held.append((mark, *[None if mark is None or one is None else one - mark
                                 for one in lanes]))
        same = held[0] == held[1]
        print(f"  {offset:>11.2f}   {str(held[0]):>22}{str(held[1]):>22}"
              f"{'' if same else '  <<'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
