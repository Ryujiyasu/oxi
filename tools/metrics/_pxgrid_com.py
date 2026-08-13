# -*- coding: utf-8 -*-
"""Are Word's paragraph tops on the 96dpi device-pixel grid (0.75pt)?

policies__000f7115 and reports__00377a16 both place EVERY measured paragraph
at an exact multiple of 0.75pt, and their line pitches are exactly what a
cumulative round-to-pixel of the true multiplied line height produces
(Arial 12 x1.15 = 15.87 -> 21px = 15.75 per line, but 4 lines = 85px = 63.75;
TNR 12 x1.5 = 20.70 -> 7 lines = 193px = 144.75).  This script tests the grid
claim across whole documents rather than at one boundary.

  python _pxgrid_com.py <docx> [<docx> ...]
"""
import os
import sys
from collections import Counter

import win32com.client as w

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

QUANTA = [0.75, 0.5, 0.25, 0.05]


def check(app, path: str) -> None:
    d = app.Documents.Open(os.path.abspath(path), ReadOnly=True)
    try:
        d.Repaginate()
        n = d.Paragraphs.Count
        ys = []
        for i in range(1, n + 1):
            rng = d.Paragraphs(i).Range
            c = d.Range(rng.Start, rng.Start)
            ys.append(round(c.Information(6), 2))
        hits = Counter()
        for y in ys:
            for q in QUANTA:
                if abs(y / q - round(y / q)) < 1e-6:
                    hits[q] += 1
        print(f"{os.path.basename(path):28s} n={len(ys):5d}  " +
              "  ".join(f"{q}pt {hits[q] * 100.0 / max(1, len(ys)):5.1f}%" for q in QUANTA))
        off = [y for y in ys if abs(y / 0.75 - round(y / 0.75)) >= 1e-6]
        if off:
            print(f"    off-grid sample: {off[:12]}")
    finally:
        d.Close(False)


def main() -> None:
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    try:
        for p in sys.argv[1:]:
            check(app, p)
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
