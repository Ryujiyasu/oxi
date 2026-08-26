# -*- coding: utf-8 -*-
r"""How long is a dash, and where does the first one start?

`glossary_05` is the corpus floor, and its worst rows are not text at all —
they are the dashed outlines of its flowchart boxes. Excel draws that dash
five pixels on and two off; we draw four and three, which is what DrawingML
says a `dash` is (4 and 3, in multiples of the line's width). Every dash after
the first is then in the wrong place, so a border that looks right to the eye
scores as if the whole line were missing.

So this asks Excel directly, one line-shape an arm: every preset dash the
format defines, at several widths, drawn long enough to read the pattern off
the picture. What comes back is the on/off run lengths and where the first run
begins — the phase matters as much as the ratio, since a dash a pixel out is
as wrong as a dash of the wrong length.

    python tools\metrics\_xlsx_shape_dash.py
    python tools\metrics\_xlsx_shape_dash.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_dash")

# Excel's own dash numbering, and what each one SAVES as — read back out of
# the workbook rather than guessed at. The numbers are not in the format's
# order, and two of them are the same preset told apart only by the cap:
# `dot cap="sq"` and `dot cap="rnd"` are different lines on the page.
DASHES = [
    ("solid", 1),
    ("dot cap=sq", 2),
    ("dot cap=rnd", 3),
    ("dash", 4),
    ("dashDot", 5),
    ("sysDashDotDot", 6),
    ("lgDash", 7),
    ("lgDashDot", 8),
    ("lgDashDotDot", 9),
    ("sysDash", 10),
    ("sysDot", 11),
]
WIDTHS = [0.75, 1.5, 2.25, 3.0, 3.75, 4.5]
ARMS = [(name, style, width) for name, style in DASHES for width in WIDTHS]
LONG = 480.0   # points across, plenty to read a pattern from
STEP = 24.0    # points between one arm's line and the next


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:Z200").Interior.Color = 0xFFFFFF
        for at, (_name, style, width) in enumerate(ARMS):
            top = 12.0 + at * STEP
            line = sheet.Shapes.AddLine(12.0, top, 12.0 + LONG, top)
            line.Line.ForeColor.RGB = 0
            line.Line.Weight = width
            line.Line.DashStyle = style
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range("A1:AA200").CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.9)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def pattern(dark: np.ndarray, y: int, x0: int, x1: int) -> tuple[int, list[int]] | None:
    """Where the ink starts, and the run lengths after it."""
    row = dark[y]
    lit = [x for x in range(x0, x1) if row[x]]
    if not lit:
        return None
    runs, start, last = [], lit[0], lit[0]
    for x in lit[1:]:
        if x > last + 1:
            runs.append(last - start + 1)
            runs.append(x - last - 1)
            start = x
        last = x
    runs.append(last - start + 1)
    return lit[0], runs


def rows_of(dark: np.ndarray, x0: int, x1: int) -> list[int]:
    """The y of every line drawn, taking the darkest row of each band."""
    weight = [(y, int(dark[y, x0:x1].sum())) for y in range(dark.shape[0])]
    out, band = [], []
    for y, held in weight:
        if held > 20:
            band.append((y, held))
        elif band:
            out.append(max(band, key=lambda one: one[1])[0])
            band = []
    if band:
        out.append(max(band, key=lambda one: one[1])[0])
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "dash.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    subprocess.run(
        [str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
        capture_output=True, text=True, encoding="utf-8", env=dict(os.environ),
    )
    mine = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140

    theirs = rows_of(truth, 10, truth.shape[1] - 10)
    ours = rows_of(mine, 10, mine.shape[1] - 10)
    print(f"  {len(ARMS)} arms; Excel drew {len(theirs)} line(s), we drew {len(ours)}")
    print(f"  {'dash':<16}{'pt':>5}   {'Excel: start, runs':<40}{'Oxi: start, runs'}")
    agree = 0
    for at, (name, _style, width) in enumerate(ARMS):
        if at >= len(theirs) or at >= len(ours):
            break
        one = pattern(truth, theirs[at], 5, truth.shape[1] - 5)
        two = pattern(mine, ours[at], 5, mine.shape[1] - 5)
        if one is None or two is None:
            print(f"  {name:<16}{width:>5}   nothing to read")
            continue
        same = one[1][:8] == two[1][:8]
        agree += same
        print(f"  {name:<16}{width:>5}   {str((one[0], one[1][:7])):<40}"
              f"{str((two[0], two[1][:7]))}{'' if same else '  <<'}")
    print(f"  {agree} of {min(len(ARMS), len(theirs), len(ours))} patterns match")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
