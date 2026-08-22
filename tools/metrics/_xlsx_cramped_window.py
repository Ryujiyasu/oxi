# -*- coding: utf-8 -*-
"""The same window of scanlines in both pictures, with the row edges marked.

`_xlsx_cramped_profile.py` found two lit pixels at the top of a cramped
centred row that the renderer does not draw, with a gap under them. Either
the row above spilled them or Excel's row grid is not where ours is. This
prints the raw scanlines around the case so the two can be told apart, and
where the lit pixels sit across the row so the glyph can be recognised.
"""
import sys
from pathlib import Path

import numpy as np
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))
import importlib
probe = importlib.import_module("_xlsx_valign_pixels")

WANT = ("ＭＳ Ｐゴシック", 11.0, 12.0, False)


def main():
    cases = probe.build(False, 1, 1)
    truth = np.asarray(Image.open(probe.BOOK.with_suffix(".excel.png")).convert("L"))
    ours_png, heights = probe.draw()
    ours = np.asarray(Image.open(ours_png).convert("L"))
    edges, at = {}, 0
    for index in sorted(heights):
        edges[index] = (at, at + heights[index])
        at += heights[index]
    sys.stdout.reconfigure(encoding="utf-8")
    for row, face, points, height, place, deep, bold in cases:
        if (face, points, height, bold) != WANT or place != "top":
            continue
        start = edges[row][0] - 2
        stop = edges[row + 2][1] + 2
        marks = {edges[row][0]: "top", edges[row + 1][0]: "center",
                 edges[row + 2][0]: "bottom", edges[row + 2][1]: "end"}
        print(f"{face} {points}pt rows {row}-{row+2} height {height}")
        for y in range(start, stop):
            them = np.flatnonzero(truth[y] < 128)
            mine = np.flatnonzero(ours[y] < 128)
            span = lambda lit: (f"{lit[0]:>4}-{lit[-1]:<4}({len(lit):>2})"
                                if lit.size else f"{'':>4} {'':<4}( 0)")
            print(f"  y={y:<5}{marks.get(y, ''):>7} | Excel {span(them)} "
                  f"| ours {span(mine)}")
        break


if __name__ == "__main__":
    main()
