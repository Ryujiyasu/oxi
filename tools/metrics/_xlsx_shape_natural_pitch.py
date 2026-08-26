# -*- coding: utf-8 -*-
r"""How far apart does Excel set the paragraphs of a shape's text?

Four `zuhyo` workbooks sit near the corpus floor on this. Their note is ＭＳ
明朝 11pt in three paragraphs, and Excel sets them 19 and 19 pixels apart; we
compute 19.067 — an em of 14.667 times 1.3 — and rounding each line's top from
that gives 19 then 20, so the third line lands a pixel low.

Rounding the step was tried three ways and the corpus refused all three: it
gains 0.0145 on `zuhyo` and loses 0.065 on `002`. So the question is not how
the pitch is rounded but what the pitch IS. 19 over 14.667 is 1.2955, which is
not a number anyone would write down — meaning the 1.3 is probably not the
shape of the rule at all.

This asks Excel: one text box an arm, four short paragraphs in it, swept over
faces and sizes. The answer is the distance between the ink of one paragraph
and the next, which is the pitch with nothing else in it.

    python tools\metrics\_xlsx_shape_natural_pitch.py
    python tools\metrics\_xlsx_shape_natural_pitch.py --reuse
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
SCRATCH = Path(r"C:\tmp\xlsx_shape_pitch")
FACES = ["ＭＳ 明朝", "ＭＳ ゴシック", "游ゴシック", "メイリオ", "Calibri"]
SIZES = [9.0, 10.0, 11.0, 12.0, 14.0, 18.0]
ARMS = [(face, size) for face in FACES for size in SIZES]
WIDE, TALL = 300.0, 110.0
STEP = 130.0


def build() -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:R400").Interior.Color = 0xFFFFFF
        for at, (face, size) in enumerate(ARMS):
            top = 10.0 + at * STEP
            box = sheet.Shapes.AddTextbox(1, 20.0, top, WIDE, TALL)
            box.Line.Visible = False
            box.Fill.Visible = False
            frame = box.TextFrame2
            frame.WordWrap = 0
            frame.MarginLeft = frame.MarginRight = 0
            frame.MarginTop = frame.MarginBottom = 0
            frame.TextRange.Text = "\r".join(["Xy亜", "Xy亜", "Xy亜", "Xy亜"])
            frame.TextRange.Font.Size = size
            frame.TextRange.Font.Name = face
            frame.TextRange.Font.NameFarEast = face
        book.SaveAs(str(SCRATCH / "pitch.xlsx"), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(1, 1), sheet.Cells(int(len(ARMS) * STEP / 15) + 20, 18)
                            ).CopyPicture(Appearance=1, Format=2)
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


def line_tops(dark: np.ndarray, y0: int, y1: int) -> list[int]:
    """The first row of ink of each line in this band."""
    rows = [y for y in range(y0, min(y1, dark.shape[0])) if dark[y].sum() >= 2]
    out, last = [], None
    for y in rows:
        if last is None or y > last + 1:
            out.append(y)
        last = y
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    if not args.reuse and not build():
        print("  Excel would not hand over a picture")
        return 1
    dark = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    tall_px = round(STEP * 96 / 72)
    print(f"  {'face':<14}{'pt':>5}{'em px':>8}{'x1.3':>8}   Excel's gaps")
    for at, (face, size) in enumerate(ARMS):
        tops = line_tops(dark, at * tall_px, (at + 1) * tall_px)
        gaps = [tops[i + 1] - tops[i] for i in range(len(tops) - 1)]
        em = size * 96 / 72
        print(f"  {face:<14}{size:>5}{em:>8.2f}{em * 1.3:>8.2f}   {gaps}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
