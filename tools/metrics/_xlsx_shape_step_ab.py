# -*- coding: utf-8 -*-
"""Does OUR shape text step the way Excel's does, over a hundred characters?

`_xlsx_shape_step.py` settled what Excel does: a shape steps by the DESIGN
advance — 18.667 pixels a full-width character at 14 point — in every face
measured, at a lever long enough (100 characters, thirty pixels between the
models) that the old 3-to-17-character reading could not have decided it.

This puts the renderer beside that answer on the same file: one shape per arm,
Excel's picture and ours, the same reach measured in both.

    python tools\\metrics\\_xlsx_shape_pitch_ab.py
"""

from __future__ import annotations

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
SCRATCH = Path(r"C:\tmp\xlsx_shape_step_ab")
LETTER, END = "あ", "ぬ"
LENGTHS = (10, 100)
# The last arm is `002`'s own note text, repeated, because a run of one kana
# cannot see a per-character difference in the marks and full-width digits a
# real note is made of.
NOTE = "※２・３の数量欄は、小数点第２位まで表示されます。"
ARMS = [("メイリオ", 14.0), ("メイリオ", 11.0), ("ＭＳ Ｐゴシック", 14.0),
        ("ＭＳ ゴシック", 14.0), ("游ゴシック", 14.0), ("ＭＳ 明朝", 14.0)]
UNITS = {("メイリオ", 14.0): NOTE}
TOP, HIGH = 40.0, 40.0


def build(made: Path) -> list[tuple[str, float, int, float, float]]:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    placed = []
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:CZ60").Interior.Color = 0xFFFFFF
        at = TOP
        for face, points in ARMS:
            for count in (LENGTHS if (face, points) not in UNITS else (1, 5)):
                shape = sheet.Shapes.AddShape(1, 20.0, at, 2400.0, HIGH)
                frame = shape.TextFrame2
                frame.WordWrap = False
                frame.AutoSize = 0
                frame.VerticalAnchor = 1
                unit = UNITS.get((face, points), LETTER)
                frame.TextRange.Text = unit * count + END
                frame.TextRange.Font.Size = points
                frame.TextRange.Font.Name = face
                try:
                    frame.TextRange.Font.NameFarEast = face
                except Exception:
                    pass
                frame.TextRange.Font.Bold = False
                frame.TextRange.Font.Fill.ForeColor.RGB = 0
                shape.Fill.Visible = False
                shape.Line.Visible = False
                placed.append((face, points, count, shape.Top, shape.Height))
                at += HIGH + 8.0
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range("A1:CZ60").CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.6)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                break
        else:
            return []
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return placed


def reach(picture: np.ndarray, top: float, high: float) -> int | None:
    band = picture[round(top * 96 / 72):round((top + high) * 96 / 72)]
    ink = band < 120
    if not ink.any():
        return None
    cols = np.where(ink.any(axis=0))[0]
    return int(cols.max() - cols.min())


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "pitch.xlsx"
    placed = build(made)
    if not placed:
        print("  Excel would not hand over a picture")
        return 1
    ours_png = SCRATCH / "oxi.png"
    subprocess.run([str(RENDERER), str(made), str(ours_png), "96"],
                   capture_output=True, check=False)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L"))
    ours = np.asarray(Image.open(ours_png).convert("L"))
    print("  face          size    Excel step    Oxi step    difference over 90")
    read: dict[tuple[str, float, int], tuple[int | None, int | None]] = {}
    for face, points, count, top, high in placed:
        read[(face, points, count)] = (reach(truth, top, high), reach(ours, top, high))
    for face, points in ARMS:
        lengths_here = LENGTHS if (face, points) not in UNITS else (1, 5)
        short = read.get((face, points, lengths_here[0]))
        long = read.get((face, points, lengths_here[-1]))
        if not short or not long or None in short or None in long:
            print(f"  {face} {points} — nothing to read")
            continue
        lengths = LENGTHS if (face, points) not in UNITS else (1, 5)
        span = (lengths[-1] - lengths[0]) * len(UNITS.get((face, points), LETTER))
        theirs = (long[0] - short[0]) / span
        mine = (long[1] - short[1]) / span
        print(f"  {face:<14}{points:>4.0f}  {theirs:>11.3f} {mine:>11.3f}"
              f"   {(mine - theirs) * span:>+8.1f}px")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
