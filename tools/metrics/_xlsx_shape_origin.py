# -*- coding: utf-8 -*-
"""Where does a shape's first glyph start, across the box's own edge?

Every arm of `_xlsx_shape_mixed_ab.py` read Excel's first ink one pixel to the
LEFT of ours, whatever the face — a constant, and so worth its own probe. The
text origin is the box's left edge plus the inset (`lIns`, 91440 EMU = 9.6px by
default), and both are fractions of a pixel.

Two shapes to a row: a filled one with no text, to read where the BOX lands,
and a text one at the same left, to read where the INK lands. The difference is
the inset and the glyph's own bearing, with the box's rounding cancelled. The
left is swept a quarter-point at a time so the fraction runs through a whole
pixel.

    python tools\\metrics\\_xlsx_shape_origin.py
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_origin")
LETTER = "日"
FACE, POINTS = "ＭＳ ゴシック", 14.0
LEFTS = [20.0, 20.25, 20.5, 20.75, 21.0, 21.25, 21.5, 21.75]
# `002`'s notes are oneCellAnchor (Excel writes that for a shape set to move
# with its cell but not size with it), and the twoCellAnchor arms alone cannot
# say whether the anchor kind changes where the text starts.
PLACEMENTS = [(3, "two-cell"), (2, "one-cell")]
TOP, HIGH = 30.0, 34.0


def build(made: Path) -> list[tuple[float, float, float]]:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    placed = []
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:BZ100").Interior.Color = 0xFFFFFF
        at = TOP
        for placement, _name in PLACEMENTS:
          for left in LEFTS:
              # The box: filled, no text, so its own left edge is readable.
              box = sheet.Shapes.AddShape(1, left, at, 400.0, HIGH - 6)
              box.Fill.Visible = True
              box.Fill.ForeColor.RGB = 0
              box.Line.Visible = False
              box.Placement = placement
              box.TextFrame2.TextRange.Text = ""
              # The text: same left, no fill, no line.
              words = sheet.Shapes.AddShape(1, left, at + HIGH, 400.0, HIGH - 6)
              frame = words.TextFrame2
              frame.WordWrap = False
              frame.AutoSize = 0
              frame.VerticalAnchor = 1
              frame.TextRange.Text = LETTER
              frame.TextRange.Font.Size = POINTS
              frame.TextRange.Font.Name = FACE
              try:
                  frame.TextRange.Font.NameFarEast = FACE
              except Exception:
                  pass
              frame.TextRange.Font.Fill.ForeColor.RGB = 0
              words.Fill.Visible = False
              words.Line.Visible = False
              words.Placement = placement
              # The same again, put against the RIGHT edge, so the other inset
              # can be read the same way.
              tail = sheet.Shapes.AddShape(1, left, at + HIGH * 2, 400.0, HIGH - 6)
              frame = tail.TextFrame2
              frame.WordWrap = False
              frame.AutoSize = 0
              frame.VerticalAnchor = 1
              frame.TextRange.Text = LETTER
              frame.TextRange.Font.Size = POINTS
              frame.TextRange.Font.Name = FACE
              try:
                  frame.TextRange.Font.NameFarEast = FACE
              except Exception:
                  pass
              frame.TextRange.Font.Fill.ForeColor.RGB = 0
              frame.TextRange.ParagraphFormat.Alignment = 3   # msoAlignRight
              tail.Fill.Visible = False
              tail.Line.Visible = False
              tail.Placement = placement
              placed.append((box.Top, words.Top, tail.Top))
              at += HIGH * 3
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range("A1:BZ100").CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.8)
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


def first_ink(picture: np.ndarray, top: float) -> int | None:
    band = picture[round(top * 96 / 72):round((top + HIGH - 6) * 96 / 72)]
    lit = np.where((band < 120).any(axis=0))[0]
    return int(lit.min()) if len(lit) else None


def last_ink(picture: np.ndarray, top: float) -> int | None:
    band = picture[round(top * 96 / 72):round((top + HIGH - 6) * 96 / 72)]
    lit = np.where((band < 120).any(axis=0))[0]
    return int(lit.max()) if len(lit) else None


def main() -> int:
    made = SCRATCH / "origin.xlsx"
    placed = build(made)
    if not placed:
        print("  Excel would not hand over a picture")
        return 1
    ours_png = SCRATCH / "oxi.png"
    subprocess.run([str(RENDERER), str(made), str(ours_png), "96"],
                   capture_output=True, check=False)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L"))
    mine = np.asarray(Image.open(ours_png).convert("L"))
    print(f"  {FACE} {POINTS}pt, one {LETTER!r}, default inset 9.6px")
    print("  anchor    left pt   left px    box L/R      ink-boxL: Excel Oxi    boxR-ink: Excel Oxi")
    for (left, kind), (box_top, text_top, tail_top) in zip(
            [(left, name) for _p, name in PLACEMENTS for left in LEFTS], placed):
        box_they, box_we = first_ink(truth, box_top), first_ink(mine, box_top)
        end_they, end_we = last_ink(truth, box_top), last_ink(mine, box_top)
        ink_they, ink_we = first_ink(truth, text_top), first_ink(mine, text_top)
        tail_they, tail_we = last_ink(truth, tail_top), last_ink(mine, tail_top)
        if None in (box_they, box_we, ink_they, ink_we, tail_they, tail_we):
            print(f"  {left:>7}   nothing to read")
            continue
        print(f"  {kind:<9}{left:>7} {left * 96 / 72:>9.3f}   {box_they:>4}/{end_they:<4}"
              f"      {ink_they - box_they:>5} {ink_we - box_we:>3}"
              f"          {end_they - tail_they:>5} {end_we - tail_we:>3}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
