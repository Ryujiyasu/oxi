# -*- coding: utf-8 -*-
"""Is the ink at the top of a cramped row its own, or the row above's?

`_xlsx_valign_pixels.py` reports each row band's first and last lit scanline.
A centred row too short for its line reads 0 there, which was read as "Excel
raises the line until its ink touches the row's top". The same 0 appears if
the row above simply spills a pixel of ink downward and Excel does not clip
it. The two differ in the shape of the band: a spill leaves a gap under it,
a raised line does not.

Run after `_xlsx_valign_pixels.py` has shot the book (uses its pictures).
"""
import sys
from pathlib import Path

import numpy as np
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))
import importlib
probe = importlib.import_module("_xlsx_valign_pixels")

WANTED = [("ＭＳ Ｐゴシック", 11.0, 12.0, False),
          ("ＭＳ Ｐゴシック", 14.0, 15.0, False),
          ("ＭＳ Ｐゴシック", 18.0, 18.0, True),
          ("ＭＳ Ｐゴシック", 18.0, 19.5, True),
          ("游ゴシック", 11.0, 13.5, True)]


def main():
    cases = probe.build(False, 1, 1)
    picture = probe.BOOK.with_suffix(".excel.png")
    ours_png, heights = probe.draw()
    truth = np.asarray(Image.open(picture).convert("L"))
    ours = np.asarray(Image.open(ours_png).convert("L"))
    edges, at = {}, 0
    for index in sorted(heights):
        edges[index] = (at, at + heights[index])
        at += heights[index]
    sys.stdout.reconfigure(encoding="utf-8")
    for row, face, points, height, place, deep, bold in cases:
        if (face, points, height, bold) not in WANTED:
            continue
        top, foot = edges[row]
        band_t = (truth[top:foot] < 128).sum(axis=1)
        band_o = (ours[top:foot] < 128).sum(axis=1)
        print(f"{face}{' bold' if bold else ''} {points}pt row {height} "
              f"({foot - top}px) {place}")
        print(f"   Excel {list(map(int, band_t))}")
        print(f"   ours  {list(map(int, band_o))}")


if __name__ == "__main__":
    main()
