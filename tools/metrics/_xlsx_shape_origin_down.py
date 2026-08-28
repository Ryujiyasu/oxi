# -*- coding: utf-8 -*-
r"""Where does a shape's text start DOWN the box?

`_xlsx_shape_origin.py` settled the sides: Excel adds the inset to the box's
own exact edge and puts the answer on a whole pixel once, rather than rounding
the edge and the inset separately. The vertical was never asked, and the code
still rounds the box's top first and its inset second.

Aligning the two is a symmetry argument, not a measurement — and the corpus is
neutral on it (net -0.0004, though the floor book gains 0.0013), so symmetry is
all there would be to go on. This asks Excel instead.

Each arm is a pair at the same top, swept a quarter of a pixel at a time: a
filled box with no text, whose own top edge is readable, and a text shape whose
first line of ink is what is being measured. The distance between them is the
inset as Excel applies it.

    python tools\metrics\_xlsx_shape_origin_down.py
    python tools\metrics\_xlsx_shape_origin_down.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_origin_down")
LETTER = "日"
FACE, POINTS = "ＭＳ ゴシック", 14.0
# A pixel is 0.75pt, so these step the top a quarter of a pixel at a time.
TOPS = [30.0, 30.1875, 30.375, 30.5625, 30.75, 30.9375, 31.125, 31.3125]
WIDE, TALL = 90.0, 40.0
GAP = 60.0


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:P200").Interior.Color = 0xFFFFFF
        for at, top in enumerate(TOPS):
            here = 20.0 + at * (TALL + GAP)
            # The box: filled, no text, so its own top edge is readable.
            box = sheet.Shapes.AddShape(1, 20.0, here, WIDE, TALL)
            box.Fill.Visible = True
            box.Fill.ForeColor.RGB = 0
            box.Line.Visible = False
            # The text: the same offset within its own cell, no fill, no line.
            words = sheet.Shapes.AddShape(1, 200.0, here, WIDE, TALL)
            words.Fill.Visible = False
            words.Line.Visible = False
            frame = words.TextFrame2
            frame.WordWrap = False
            frame.AutoSize = 0
            frame.VerticalAnchor = 1          # top
            frame.TextRange.Text = LETTER
            frame.TextRange.Font.Size = POINTS
            frame.TextRange.Font.Name = FACE
            frame.TextRange.Font.NameFarEast = FACE
            # Black, said out loud. A shape made with no fill keeps the text
            # colour its theme gives it, which is WHITE — Excel drew every arm
            # and the reader found no ink in any of them, while ours drew its
            # own default black and looked like the only side working.
            frame.TextRange.Font.Fill.ForeColor.RGB = 0
            # Nudge both by the same fraction of a point.
            box.Top = here + (top - 30.0)
            words.Top = here + (top - 30.0)
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(1, 1), sheet.Cells(
                    int(len(TOPS) * (TALL + GAP) / 15) + 12, 16)).CopyPicture(
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
    made = SCRATCH / "origin.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    subprocess.run(
        [str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
        capture_output=True, text=True, encoding="utf-8", env=dict(os.environ),
    )
    mine = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140
    band = round((TALL + GAP) * 96 / 72)
    print(f"  {'top pt':>9}{'px':>8}   {'Excel box/ink/inset':>22}{'Oxi box/ink/inset':>22}")
    agree = 0
    for at, top in enumerate(TOPS):
        y0, y1 = at * band, (at + 1) * band
        held = []
        for dark in (truth, mine):
            box = first_ink(dark, y0, y1, 30, 140)
            ink = first_ink(dark, y0, y1, 270, 400)
            held.append((box, ink, None if box is None or ink is None else ink - box))
        same = held[0] == held[1]
        agree += same
        print(f"  {top:>9}{top * 96 / 72:>8.2f}   {str(held[0]):>22}{str(held[1]):>22}"
              f"{'' if same else '  <<'}")
    print(f"  {agree} of {len(TOPS)} arms agree")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
