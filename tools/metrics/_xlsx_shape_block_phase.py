# -*- coding: utf-8 -*-
r"""Does a shape's block of lines start on a whole pixel, or where it lands?

The renderer rounds the block's start and walks the exact pitch from there:

    top of line i = round(area top) + round(i x pitch + leading)

`_xlsx_shape_pitch_size.py` has since shown the pitch itself is right to a
hundredth of a pixel over four faces and eight sizes, so what is left on
`tb_r8_jizensoudan` — its third line a pixel low — has to be the START.

This asks Excel without needing to know either the inset or the leading. Sweep
the box's top an eighth of a pixel at a time and read where every line of a
tall block lands. Then look at the offsets WITHIN the block, each line against
the first:

  * if the start is put on a whole pixel first, those offsets are round(i x
    pitch) and are the SAME at every phase — the fraction is gone before the
    lines are laid;
  * if the exact start is carried through, they are round(f + i x pitch) -
    round(f) and MOVE as the phase moves.

Eight phases x sixteen lines, and the two answers differ in dozens of places.

    python tools\metrics\_xlsx_shape_block_phase.py
    python tools\metrics\_xlsx_shape_block_phase.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_block_phase")

# A letter with ink in every row it spans, so a line reads as one run.
LETTER = "H"
LINES = 16
# Two faces whose pitch has a very different fraction: 游ゴシック 12pt steps
# 26.762 and Yu Gothic UI 11pt steps 25.360. A fraction near a half crosses a
# rounding boundary every other line; one near zero almost never does, and a
# rule that only holds for the first is not a rule.
PANELS = [("游ゴシック", 12.0), ("Yu Gothic UI", 11.0)]
# A pixel is 0.75pt, so these step the top an eighth of a pixel at a time and
# cover exactly one whole pixel.
TOPS = [30.0 + step * 0.09375 for step in range(8)]
WIDE = 90.0
TALL = 380.0            # room for sixteen lines at the largest pitch
GAP = 30.0
LANE = 180.0            # how far apart the panels stand


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:Z4000").Interior.Color = 0xFFFFFF
        for at, top in enumerate(TOPS):
            here = 20.0 + at * (TALL + GAP)
            # The box: filled, no text, so its own top edge is readable and
            # the first line can be held against it.
            box = sheet.Shapes.AddShape(1, 20.0, here, WIDE, 20.0)
            box.Fill.Visible = True
            box.Fill.ForeColor.RGB = 0
            box.Line.Visible = False
            box.Top = here + (top - 30.0)
            for panel, (face, points) in enumerate(PANELS):
                words = sheet.Shapes.AddShape(
                    1, 160.0 + panel * LANE, here, WIDE, TALL)
                words.Fill.Visible = False
                words.Line.Visible = False
                frame = words.TextFrame2
                frame.WordWrap = False
                frame.AutoSize = 0
                frame.VerticalAnchor = 1          # top
                frame.TextRange.Text = chr(13).join([LETTER] * LINES)
                frame.TextRange.Font.Size = points
                frame.TextRange.Font.Name = face
                frame.TextRange.Font.NameFarEast = face
                # A shape made with no fill keeps the text colour its theme
                # gives it, which is WHITE: say black out loud or Excel draws
                # nothing and the reader calls it a disagreement.
                frame.TextRange.Font.Fill.ForeColor.RGB = 0
                words.Top = here + (top - 30.0)
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(1, 1), sheet.Cells(
                    int(len(TOPS) * (TALL + GAP) / 15) + 12, 30)).CopyPicture(
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


def runs(dark: np.ndarray, y0: int, y1: int, x0: int, x1: int) -> list[int]:
    """The first row of ink of each line in a window."""
    lit = dark[y0:min(y1, dark.shape[0]), x0:x1].sum(axis=1) >= 2
    found, running = [], False
    for step, on in enumerate(lit):
        if on and not running:
            found.append(y0 + step)
        running = bool(on)
    return found


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "blockphase.xlsx"
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

    for panel, (face, points) in enumerate(PANELS):
        x0 = round((160.0 + panel * LANE) * 96 / 72) - 8
        x1 = x0 + round(WIDE * 96 / 72)
        print(f"\n{face} {points}pt — offsets within the block, line 0 held at 0")
        print(f"  {'top pt':>9}{'phase':>7}{'box':>6}{'1st':>6}  {'Excel':<58}{'Oxi'}")
        seen = {}
        for at, top in enumerate(TOPS):
            y0, y1 = at * band, (at + 1) * band
            held = []
            for dark in (truth, mine):
                box = runs(dark, y0, y1, 30, 140)
                lines = runs(dark, y0, y1, x0, x1)
                held.append((box[0] if box else None, lines))
            said = []
            for box, lines in held:
                if len(lines) < 2:
                    said.append(("-", "-"))
                    continue
                said.append((
                    "" if box is None else str(lines[0] - box),
                    " ".join(str(one - lines[0]) for one in lines[1:]),
                ))
            seen.setdefault(said[0][1], []).append(at)
            same = said[0] == said[1]
            print(f"  {top:>9.5f}{(top * 96 / 72) % 1.0:>7.3f}"
                  f"{held[0][0] if held[0][0] is not None else -1:>6}{said[0][0]:>6}"
                  f"  {said[0][1][:56]:<58}{said[1][1][:56]}"
                  f"{'' if same else '  <<'}")
        # The whole question in one line: how many DIFFERENT patterns Excel
        # gave over the eight phases. One means the start is rounded before
        # the lines are laid; more means the fraction survives into them.
        print(f"  Excel gave {len(seen)} distinct pattern(s) over {len(TOPS)} phases")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
