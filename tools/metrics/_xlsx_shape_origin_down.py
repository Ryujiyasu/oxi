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
# The face and size the one corpus shape that moved is set in: its theme names
# 游ゴシック for Japanese and the runs say 12 point. A rule settled on one
# face is not settled until the face that disagreed has been asked too.
FACE2, POINTS2 = "游ゴシック", 12.0
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
            # And the shape the corpus actually holds: text that WRAPS, more
            # lines than the box has room for, and overflow clipped. The
            # single-line arm above settled the rounding; `tb_r8_jizensoudan`
            # then went the other way, and its shape differs in exactly these
            # three things.
            many = sheet.Shapes.AddShape(1, 380.0, here, WIDE, TALL)
            many.Fill.Visible = False
            many.Line.Visible = False
            frame = many.TextFrame2
            frame.WordWrap = True
            frame.AutoSize = 0
            frame.VerticalAnchor = 1          # top
            frame.TextRange.Text = (LETTER * 6 + chr(13)) * 4
            frame.TextRange.Font.Size = POINTS
            frame.TextRange.Font.Name = FACE
            frame.TextRange.Font.NameFarEast = FACE
            frame.TextRange.Font.Fill.ForeColor.RGB = 0
            other = sheet.Shapes.AddShape(1, 560.0, here, WIDE, TALL)
            other.Fill.Visible = False
            other.Line.Visible = False
            frame = other.TextFrame2
            frame.WordWrap = True
            frame.AutoSize = 0
            frame.VerticalAnchor = 1          # top
            frame.TextRange.Text = (LETTER * 6 + chr(13)) * 4
            frame.TextRange.Font.Size = POINTS2
            frame.TextRange.Font.Name = FACE2
            frame.TextRange.Font.NameFarEast = FACE2
            frame.TextRange.Font.Fill.ForeColor.RGB = 0
            # And one with a BORDER. Every shape above is drawn with no line at
            # all, and the corpus shape that disagreed carries `w="28575"` —
            # 2.25 points, three pixels, straddling the box's own edge. A rule
            # settled on borderless shapes says nothing about that.
            ruled = sheet.Shapes.AddShape(1, 740.0, here, WIDE, TALL)
            ruled.Fill.Visible = False
            ruled.Line.Visible = True
            ruled.Line.Weight = 2.25
            ruled.Line.ForeColor.RGB = 0
            frame = ruled.TextFrame2
            frame.WordWrap = True
            frame.AutoSize = 0
            frame.VerticalAnchor = 1          # top
            frame.TextRange.Text = (LETTER * 6 + chr(13)) * 4
            frame.TextRange.Font.Size = POINTS2
            frame.TextRange.Font.Name = FACE2
            frame.TextRange.Font.NameFarEast = FACE2
            frame.TextRange.Font.Fill.ForeColor.RGB = 0
            # Nudge all five by the same fraction of a point.
            box.Top = here + (top - 30.0)
            words.Top = here + (top - 30.0)
            many.Top = here + (top - 30.0)
            other.Top = here + (top - 30.0)
            ruled.Top = here + (top - 30.0)
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(1, 1), sheet.Cells(
                    int(len(TOPS) * (TALL + GAP) / 15) + 12, 34)).CopyPicture(
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
    print(f"  {'top pt':>9}{'px':>8}   {'Excel box/one/wrap/oth/ruled':>30}{'Oxi box/one/wrap/oth/ruled':>30}")
    agree = 0
    for at, top in enumerate(TOPS):
        y0, y1 = at * band, (at + 1) * band
        held = []
        for dark in (truth, mine):
            box = first_ink(dark, y0, y1, 30, 140)
            one = first_ink(dark, y0, y1, 270, 400)
            wrapped = first_ink(dark, y0, y1, 510, 640)
            second = first_ink(dark, y0, y1, 750, 880)
            # Inside the bordered box, past its own top rule: what is wanted is
            # the text, and a 2.25pt rule straddling the edge answers first —
            # it reaches ABOVE the box top, so the search starts below it. The
            # window stays clear of the box's SIDE rules as well: they put ink on
            # every row inside it and the reader answers `box + 6` for every arm.
            boxed = None if box is None else first_ink(dark, box + 6, y1, 1010, 1090)
            held.append((
                box,
                None if box is None or one is None else one - box,
                None if box is None or wrapped is None else wrapped - box,
                None if box is None or second is None else second - box,
                None if box is None or boxed is None else boxed - box,
            ))
        same = held[0] == held[1]
        agree += same
        print(f"  {top:>9}{top * 96 / 72:>8.2f}   {str(held[0]):>30}{str(held[1]):>30}"
              f"{'' if same else '  <<'}")
    print(f"  {agree} of {len(TOPS)} arms agree")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
