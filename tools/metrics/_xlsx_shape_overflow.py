"""Where a shape puts a block of text that does not fit its box.

A shape anchored `ctr` centres its text in its box. `sanko_tool`'s roundRect
holds ten lines in a box eight lines deep, and Oxi's centring puts the block a
whole line pitch above where Excel draws it — so the rule Excel uses when the
block is TALLER than the box is not the rule it uses when it fits.

The arm: one shape per box height, the same text every time, `anchor` swept
over t/ctr/b. Excel draws it, the sheet is copied as a picture, and the ink
bands are read off the bitmap. The box shrinks from comfortably taller than
the block to half its height, so the answer is read on both sides of the
crossing rather than at one point.

Two traps this walked into and holds fixed:
  * a ruled, rounded box is what a scanline finds first — the outline has to
    go, not just the fill;
  * a shape Excel adds wears the theme's style, whose text is WHITE, so with
    the fill gone the words disappear instead of appearing.

Run: python tools/metrics/_xlsx_shape_overflow.py
"""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

HERE = Path(__file__).resolve().parent
OUT = HERE.parent.parent / "pipeline_data" / "shape_overflow"

TEXT = ["いろはにほへと", "ちりぬるを", "わかよたれそ", "つねならむ", "うゐのおくやま"]
FACE = "ＭＳ ゴシック"
SIZE = 12.0
LEFT_PT = 200.0
TOP_PT = 100.0
WIDE_PT = 300.0
ANCHORS = {"t": 1, "ctr": 3, "b": 4}
# 1 = rect (no corner inset), 5 = roundRect (text rect pulled in by the
# rounding, so the two together separate the inset from the line count).
PRESETS = {"rect": 1, "roundRect": 5}


def ink_rows(image: Image.Image, left: int, right: int) -> list[tuple[int, int]]:
    """The bands of lit rows in a column range, top to bottom."""
    grey = np.asarray(image.convert("L"))[:, left:right]
    lit = (grey < 120).sum(axis=1)
    bands: list[tuple[int, int]] = []
    start = None
    for row, count in enumerate(lit):
        if count > 1 and start is None:
            start = row
        if count <= 1 and start is not None:
            bands.append((start, row - 1))
            start = None
    if start is not None:
        bands.append((start, len(lit) - 1))
    return bands


def picture(sheet) -> Image.Image | None:
    """The sheet as a bitmap. Excel refuses CopyPicture now and then; ask again."""
    for _ in range(6):
        try:
            sheet.Activate()
            sheet.Range("A1:Z60").CopyPicture(Appearance=1, Format=2)
        except Exception:
            time.sleep(0.6)
            continue
        time.sleep(0.4)
        held = ImageGrab.grabclipboard()
        if held is not None:
            return held
        time.sleep(0.4)
    return None


def measure(sheet, height_pt: float, anchor: str, lines: int, preset: str) -> dict:
    shape = sheet.Shapes.AddShape(PRESETS[preset], LEFT_PT, TOP_PT, WIDE_PT, height_pt)
    try:
        frame = shape.TextFrame2
        frame.WordWrap = True
        frame.AutoSize = 0
        frame.VerticalAnchor = ANCHORS[anchor]
        frame.TextRange.Text = "\n".join(TEXT[:lines])
        frame.TextRange.Font.Size = SIZE
        frame.TextRange.Font.Name = FACE
        frame.TextRange.Font.Fill.ForeColor.RGB = 0
        shape.Fill.Visible = False
        shape.Line.Visible = False
        held = picture(sheet)
        top = round(TOP_PT * 96 / 72)
        said = {
            "height": height_pt,
            "anchor": anchor,
            "lines": lines,
            "preset": preset,
            "box_top": top,
            "box_bottom": top + round(height_pt * 96 / 72),
            "bands": None,
        }
        if held is not None:
            left = round(LEFT_PT * 96 / 72) + 6
            right = round((LEFT_PT + WIDE_PT) * 96 / 72) - 6
            said["bands"] = ink_rows(held, left, right)
        return said
    finally:
        shape.Delete()


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    held = []
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:Z60").Interior.Color = 0xFFFFFF
        for preset in ("rect", "roundRect"):
            for height_pt in (140.0, 120.0, 100.0, 90.0, 80.0, 70.0, 60.0, 50.0, 40.0):
                for anchor in ("t", "ctr", "b"):
                    lines = 5
                    said = measure(sheet, height_pt, anchor, lines, preset)
                    held.append(said)
                    bands = said["bands"] or []
                    tops = [band[0] for band in bands]
                    print(
                        f"{preset:<9} height {height_pt:5.1f} anchor {anchor:<3}"
                        f" box {said['box_top']}..{said['box_bottom']}"
                        f" n {len(bands)} tops {tops}"
                    )
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    (OUT / "_overflow.json").write_text(json.dumps(held, indent=1), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
