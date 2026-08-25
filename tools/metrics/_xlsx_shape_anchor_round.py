# -*- coding: utf-8 -*-
r"""Which way does Excel round the leftover when it centres a shape's text?

`dc4fcff7f5f8_001` — the second-lowest workbook — opens with a text box over a
115.5pt row, `anchor="ctr"`, five lines of mixed sizes. Every one of those lines
sits ONE pixel high in our picture, and the pitch between them is right to the
pixel, so what is out is where the block starts: the leftover above it.

A centred block starts at `(box - block) / 2` and that halving has to fall
somewhere. This walks the box's height a quarter-point at a time so the
leftover runs through odd and even, and reads where the ink starts from the
box's own top — a filled twin with no text beside every arm gives that top
without trusting our own anchor arithmetic.

Two line counts, because a block of one line and a block of three round the
same leftover from different heights.

    python tools\metrics\_xlsx_shape_anchor_round.py
    python tools\metrics\_xlsx_shape_anchor_round.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_anchor")
FACE, POINTS = "ＭＳ Ｐゴシック", 12.0
LEFT, WIDE = 24.0, 300.0
# The box's height, a quarter point at a time: at 96dpi a quarter point is a
# third of a pixel, so four steps walk a whole pixel of leftover.
HIGHS = [40.0 + step / 4.0 for step in range(12)]
LINES = [1, 3]
GAP = 36.0


def build(made: Path) -> list[tuple[int, float, float]]:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    placed = []
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:P200").Interior.Color = 0xFFFFFF
        at = 24.0
        for count in LINES:
            for high in HIGHS:
                # One shape an arm, filled PALE and lettered BLACK: the box's
                # own top is the first row that is not white, the text's is the
                # first row that is dark, and neither reading has to trust our
                # own arithmetic for where the box lands.
                words = sheet.Shapes.AddShape(1, LEFT, at, WIDE, high)
                words.Fill.Visible = True
                words.Fill.ForeColor.RGB = 0xE0E0E0
                words.Line.Visible = False
                # `AddShape` hands back the theme's style, and that carries a
                # shadow: it is ink below and right of the box, and it would be
                # read as the box's own foot.
                try:
                    words.Shadow.Visible = False
                except Exception:
                    pass
                frame = words.TextFrame2
                frame.WordWrap = False
                frame.AutoSize = 0
                frame.VerticalAnchor = 3            # msoAnchorMiddle
                frame.MarginTop = 0
                frame.MarginBottom = 0
                frame.TextRange.Text = "\r".join(["国国国"] * count)
                frame.TextRange.Font.Size = POINTS
                frame.TextRange.Font.Name = FACE
                try:
                    frame.TextRange.Font.NameFarEast = FACE
                except Exception:
                    pass
                frame.TextRange.Font.Fill.ForeColor.RGB = 0
                placed.append((count, words.Top, words.Top))
                at += high + GAP
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range("A1:P200").CopyPicture(Appearance=1, Format=2)
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


def edge(picture: np.ndarray, top: float, high: float) -> tuple[int, int, int, int] | None:
    """The shape's own first and last row, and its text's, in one reading.

    The band is generous at both ends: what is being read is where the two
    edges sit relative to each other, so the window only has to hold them.
    """
    start = max(0, round(top * 96 / 72) - 6)
    band = picture[start:round((top + high) * 96 / 72) + 6]
    # The box is found by its own pale value, not by "anything that is not
    # white": Excel's picture carries the sheet's gridlines, and a gridline
    # crossing the band reads as the box's own edge. Inside the box the fill
    # covers them, so the test is the colour and not the darkness.
    paint = np.where((np.abs(band - 224) <= 6).sum(axis=1) > 40)[0]
    letters = np.where((band < 120).any(axis=1))[0]
    if not len(paint) or not len(letters):
        return None
    return (int(paint.min()) + start, int(paint.max()) + start,
            int(letters.min()) + start, int(letters.max()) + start)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "anchor.xlsx"
    if args.reuse:
        placed = []
        at = 24.0
        for count in LINES:
            for high in HIGHS:
                placed.append((count, at, at))
                at += high + GAP
    else:
        placed = build(made)
        if not placed:
            print("  Excel would not hand over a picture")
            return 1
    subprocess.run([str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
                   env={**os.environ, "OXI_XLSX_SHAPE_TEXT": "1"},
                   capture_output=True, check=False)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L"))
    mine = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L"))
    print(f"  {FACE} {POINTS}pt centred in a box, margins nil")
    print("  lines  box pt   box px   leftover  |  above: Excel Oxi   below: Excel Oxi")
    for (count, top, _same), high in zip(
            placed, [high for _count in LINES for high in HIGHS]):
        theirs = edge(truth, top, high)
        ours = edge(mine, top, high)
        if theirs is None or ours is None:
            print(f"  {count:>5}  {high:>6.2f}   nothing to read")
            continue
        above, below = theirs[2] - theirs[0], theirs[1] - theirs[3]
        we_above, we_below = ours[2] - ours[0], ours[1] - ours[3]
        print(f"  {count:>5}  {high:>6.2f}  {theirs[1] - theirs[0] + 1:>3}/{ours[1] - ours[0] + 1:<3}"
              f"  {above + below:>9} |  {above:>10} {we_above:>3}"
              f"   {below:>12} {we_below:>3}"
              f"  {'' if above == we_above else '<<'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
