# -*- coding: utf-8 -*-
r"""Where does the `cas-` title lose its four pixels?

Ten `cas-r*` workbooks open with 「内閣所管（合算）」 in a shape, asked for in
`ＤＦ特太ゴシック体` — a face this machine has not got, so SX54 answers 游ゴシック.
Our kanji land on Excel's step for step and the line still comes out four
pixels wide, which centres it two pixels left. SX100's device-advance rule does
not touch it: 游ゴシック designs every one of those characters at a whole em, so
the design and the device agree at 19.

So this reads the string itself, LEFT aligned so the centring cannot hide
anything, glyph by glyph — Excel's ink beside ours, which cancels the bearings
because they are the same glyphs. Beside it: the same string in 游ゴシック asked
for by name, and a run of kanji as a ruler. If the by-name arm steps like Excel
and the substituted one does not, the substitution is laying out differently
from the face it resolves to.

    python tools\metrics\_xlsx_shape_cas_title.py
    python tools\metrics\_xlsx_shape_cas_title.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_cas_title")
TITLE = "内閣所管（合算）"
ARMS = [
    ("ＤＦ特太ゴシック体", TITLE),
    ("游ゴシック", TITLE),
    ("ＤＦ特太ゴシック体", "内閣所管合算日"),
    ("游ゴシック", "内閣所管合算日"),
    ("ＤＦ特太ゴシック体", "日日日日日日日日"),
    ("游ゴシック", "日日日日日日日日"),
    ("ＤＦ特太ゴシック体", "（（（（"),
    ("游ゴシック", "（（（（"),
]
POINTS = 14.0
WIDE, HIGH = 260.0, 26.0
GAP = 6.0


def build(made: Path) -> list[tuple[str, str, float]]:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    placed = []
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:J60").Interior.Color = 0xFFFFFF
        at = 12.0
        for face, words in ARMS:
            shape = sheet.Shapes.AddShape(1, 18.0, at, WIDE, HIGH)
            frame = shape.TextFrame2
            frame.WordWrap = False
            frame.AutoSize = 0
            frame.VerticalAnchor = 1
            frame.TextRange.Text = words
            frame.TextRange.Font.Size = POINTS
            frame.TextRange.Font.Name = face
            try:
                frame.TextRange.Font.NameFarEast = face
            except Exception:
                pass
            frame.TextRange.Font.Fill.ForeColor.RGB = 0
            frame.TextRange.ParagraphFormat.Alignment = 1      # left
            shape.Fill.Visible = False
            shape.Line.Visible = False
            try:
                shape.Shadow.Visible = False
            except Exception:
                pass
            placed.append((face, words, shape.Top))
            at += HIGH + GAP
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range("A1:J60").CopyPicture(Appearance=1, Format=2)
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


def starts(picture: np.ndarray, top: float) -> list[int]:
    band = picture[round(top * 96 / 72):round((top + HIGH) * 96 / 72)] < 120
    col = band.any(axis=0)
    out, was = [], False
    for i, v in enumerate(col):
        if v and not was:
            out.append(i)
        was = v
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "castitle.xlsx"
    if args.reuse:
        placed, at = [], 12.0
        for face, words in ARMS:
            placed.append((face, words, at))
            at += HIGH + GAP
    else:
        placed = build(made)
        if not placed:
            print("  Excel would not hand over a picture")
            return 1
    subprocess.run([str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
                   env={**os.environ}, capture_output=True, check=False)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L"))
    mine = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L"))
    print(f"  {POINTS}pt, left aligned; the ink runs of each arm")
    for face, words, top in placed:
        theirs, ours = starts(truth, top), starts(mine, top)
        span_they = theirs[-1] - theirs[0] if len(theirs) > 1 else 0
        span_we = ours[-1] - ours[0] if len(ours) > 1 else 0
        print(f"  {face:<14}{words!r}")
        print(f"      Excel {theirs}  span {span_they}")
        print(f"      Oxi   {ours}  span {span_we}"
              f"  {'' if theirs == ours else '<<'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
