# -*- coding: utf-8 -*-
"""Is what Oxi draws simply offset from what Excel draws, and by how much?

SSIM falls hard on a page whose ink is a pixel out of place, and the drop
looks the same whether the cause is a shift or a shape. Sliding one picture
over the other tells them apart: if some offset makes the two agree, the
error is position; if none does, it is the glyphs themselves.
"""
import sys
from pathlib import Path

import numpy as np
from PIL import Image

Image.MAX_IMAGE_PIXELS = None
DIFFS = Path(r"pipeline_data\xlsx_diff_corpus")


def ink(path, box=None):
    image = Image.open(path).convert("L")
    if box:
        image = image.crop(box)
    return 255 - np.asarray(image, dtype=np.int16)


def main():
    stems = sys.argv[1:] or ["24d76e2a8663_h2daa202505_jikei"]
    for stem in stems:
        excel = ink(DIFFS / f"{stem}.excel.png")
        oxi = ink(DIFFS / f"{stem}.oxi.png")
        height = min(excel.shape[0], oxi.shape[0])
        width = min(excel.shape[1], oxi.shape[1])
        excel = excel[:height, :width]
        oxi = oxi[:height, :width]
        print(f"== {stem[:40]}  {width}x{height}")
        best = None
        for dy in range(-3, 4):
            row = []
            for dx in range(-3, 4):
                a = excel[max(0, dy):height + min(0, dy), max(0, dx):width + min(0, dx)]
                b = oxi[max(0, -dy):height + min(0, -dy), max(0, -dx):width + min(0, -dx)]
                score = float(np.abs(a.astype(np.int32) - b.astype(np.int32)).mean())
                row.append(score)
                if best is None or score < best[0]:
                    best = (score, dx, dy)
            print("   dy %+d: %s" % (dy, " ".join(f"{value:6.2f}" for value in row)))
        print(f"   best at dx {best[1]:+d} dy {best[2]:+d}, mean |difference| {best[0]:.2f}"
              f" (against {float(np.abs(excel.astype(np.int32) - oxi.astype(np.int32)).mean()):.2f} unshifted)")
        print(f"   ink: Excel {float((excel > 40).mean()) * 100:.2f}% of pixels, "
              f"Oxi {float((oxi > 40).mean()) * 100:.2f}%")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
