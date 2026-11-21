# -*- coding: utf-8 -*-
"""Drill buttons in Excel's pictures, found by their exact frame colours.

A compact-form button is a square: top and left edges #aab5c9, right and
bottom #8a9ebf, corners cut. For each Excel picture under <dir>/diff this
lists every button: left x, top y, size, in reading order -- the first ones
are Region (level 0) and Product (level 1) buttons of q01's rows.

    python tools/metrics/_xlsx_pivot_drill_boxes.py <dir>
"""
import glob
import json
import os
import sys

import numpy as np
from PIL import Image

TOP = np.array([0xAA, 0xB5, 0xC9])
BOTTOM = np.array([0x8A, 0x9E, 0xBF])


def boxes(path):
    img = np.asarray(Image.open(path).convert("RGB")).astype(int)
    top = (img == TOP).all(axis=2)
    bottom = (img == BOTTOM).all(axis=2)
    found = []
    h, w = top.shape
    for y in range(h - 2):
        xs = np.where(top[y, : min(w, 200)])[0]
        # a run of the top edge, with the left edge under its first pixel - 1
        runs = []
        for x in xs:
            if runs and x == runs[-1][1] + 1:
                runs[-1][1] = x
            else:
                runs.append([x, x])
        for a, b in runs:
            if b - a < 3 or a == 0 or not top[y + 1, a - 1]:
                continue
            size = b - a + 3  # edges are inset one pixel at the corners
            if y + size - 1 < h and bottom[y + size - 1, a:b + 1].all():
                found.append((int(a - 1), int(y), int(size)))
    return found


def main():
    root = sys.argv[1]
    out = {}
    for path in sorted(glob.glob(os.path.join(root, "diff", "*.excel.png"))):
        arm = os.path.basename(path)[: -len(".excel.png")]
        out[arm] = boxes(path)[:6]
        print(arm, out[arm])
    with open(os.path.join(root, "drill_boxes.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, indent=1)


if __name__ == "__main__":
    main()
