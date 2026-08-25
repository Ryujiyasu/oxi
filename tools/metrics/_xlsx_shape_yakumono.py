# -*- coding: utf-8 -*-
r"""What advance does a shape give a full-width MARK?

`cas-r*`'s title — 「内閣所管（合算）」, 游ゴシック 14pt through a missing face —
steps its kanji exactly as Excel does and still comes out four pixels wide over
eight characters, which centres it two pixels left. The phase model is not the
cause (SX99: it agrees on 24 of 25 arms, this very face among them), and what
is left in the string that a run of kanji has not got is the round brackets.

So each arm is 「日X日」 — the mark between two kanji, whose side bearings are
identical and cancel. The distance between the two kanji's ink is
`advance(日) + advance(X)`, and `advance(日)` comes from the 「日日」 arm, so the
mark's own advance falls out. Ours is read the same way from the same file.

    python tools\metrics\_xlsx_shape_yakumono.py
    python tools\metrics\_xlsx_shape_yakumono.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_yakumono")
FACES = [("游ゴシック", 14.0), ("ＭＳ Ｐゴシック", 14.0), ("ＭＳ ゴシック", 14.0),
         # Sizes where the DEVICE advance and `ceil(design)` part company, so
         # the two readings of Excel's step can be told apart: at 10pt ＭＳ Ｐ
         # ゴシック's ー is 12.81 design (ceil 13, device 12) and its Ａ is 9.53
         # (ceil 10, device 9).
         ("ＭＳ Ｐゴシック", 10.0), ("ＭＳ Ｐゴシック", 16.0)]
# The mark under test, between two kanji. The first arm is the ruler.
MARKS = ["日", "（", "）", "、", "。", "「", "」", "・", "ー", "，", "．", "％",
         "Ａ", "ａ", "A", "a"]
WIDE, HIGH = 200.0, 26.0
GAP = 6.0


def build(made: Path) -> list[tuple[str, float, str, float]]:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    placed = []
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:H200").Interior.Color = 0xFFFFFF
        at = 12.0
        for face, points in FACES:
            for mark in MARKS:
                words = sheet.Shapes.AddShape(1, 18.0, at, WIDE, HIGH)
                frame = words.TextFrame2
                frame.WordWrap = False
                frame.AutoSize = 0
                frame.VerticalAnchor = 1
                frame.TextRange.Text = f"日{mark}日"
                frame.TextRange.Font.Size = points
                frame.TextRange.Font.Name = face
                try:
                    frame.TextRange.Font.NameFarEast = face
                except Exception:
                    pass
                frame.TextRange.Font.Fill.ForeColor.RGB = 0
                frame.TextRange.ParagraphFormat.Alignment = 1   # left
                words.Fill.Visible = False
                words.Line.Visible = False
                try:
                    words.Shadow.Visible = False
                except Exception:
                    pass
                placed.append((face, points, mark, words.Top))
                at += HIGH + GAP
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range("A1:H200").CopyPicture(Appearance=1, Format=2)
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


def kanji_starts(picture: np.ndarray, top: float) -> tuple[int, int] | None:
    """Where the two 日 begin, in this arm's own band.

    A 日 is a solid box of ink and every mark is thinner, so the two widest
    runs of the band are the kanji — which keeps a comma or a full stop from
    being read as one of them.
    """
    start = round(top * 96 / 72)
    band = picture[start:start + round(HIGH * 96 / 72)] < 120
    col = band.any(axis=0)
    runs, at = [], None
    for i, v in enumerate(col):
        if v and at is None:
            at = i
        elif not v and at is not None:
            runs.append((at, i - at))
            at = None
    if at is not None:
        runs.append((at, len(col) - at))
    if not runs:
        return None
    # The two 日 are the widest runs of the band, and every mark and letter
    # under test is thinner: taking everything within a pixel of the widest
    # keeps a wide Latin capital from being read as one of them.
    widest = max(wide for _where, wide in runs)
    tall = [(where, wide) for where, wide in runs if wide >= widest - 1]
    if len(tall) < 2:
        return None
    return tall[0][0], tall[-1][0]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "yakumono.xlsx"
    if args.reuse:
        placed, at = [], 12.0
        for face, points in FACES:
            for mark in MARKS:
                placed.append((face, points, mark, at))
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
    print("  face              size  mark |  日+mark: Excel Oxi  |  mark alone: Excel Oxi")
    ruler = {}
    for face, points, mark, top in placed:
        theirs = kanji_starts(truth, top)
        ours = kanji_starts(mine, top)
        if theirs is None or ours is None:
            print(f"  {face:<16}{points:>5}  {mark}  | nothing to read")
            continue
        span_they, span_we = theirs[1] - theirs[0], ours[1] - ours[0]
        if mark == "日":
            ruler[(face, points)] = (span_they // 2, span_we // 2)
        one = ruler.get((face, points))
        note = ""
        if one:
            note = (f"  |  {span_they - one[0]:>10} {span_we - one[1]:>3}"
                    f" {'' if span_they - one[0] == span_we - one[1] else '<<'}")
        print(f"  {face:<16}{points:>5}   {mark}  |  {span_they:>9} {span_we:>3}"
              f" {'' if span_they == span_we else '<<'}{note}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
