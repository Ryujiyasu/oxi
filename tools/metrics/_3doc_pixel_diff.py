"""Pixel measure top-3 horizontal lines in 3 floating-table docs.

For each: 1ec1 p.1, 459f p.1, 2ea81a p.1.
Goal: identify per-doc Y offset to inform per-doc fix.
"""
from PIL import Image
import numpy as np

DOCS = [
    ("1ec1091177b1_006", "page_0001.png"),
    ("459f05f1e877_kyodokenkyuyoushiki01", "page_0001.png"),
    ("2ea81a8441cc_0025006-192", "page_0001.png"),
]

PT_PER_PX = 72.0 / 150


def find_lines(img_path, min_dark=300):
    img = Image.open(img_path).convert("L")
    arr = np.array(img)
    counts = (arr < 128).sum(axis=1)
    h = arr.shape[0]
    lines = []
    in_run = False; run_start = 0
    for y in range(h):
        if counts[y] > min_dark:
            if not in_run:
                in_run = True; run_start = y
        else:
            if in_run:
                in_run = False
                mid = (run_start + y) // 2
                lines.append(mid)
    return lines


for doc, pg in DOCS:
    word_path = f"pipeline_data/word_png/{doc}/{pg}"
    oxi_path = f"pipeline_data/oxi_png/{doc}/{pg}"
    word_lines = find_lines(word_path)
    oxi_lines = find_lines(oxi_path)
    print(f"\n{doc}/{pg}:")
    print(f"  {'#':>3} {'Word_pt':>8} {'Oxi_pt':>8} {'delta':>7}")
    for i in range(min(8, len(word_lines), len(oxi_lines))):
        w_pt = word_lines[i] * PT_PER_PX
        o_pt = oxi_lines[i] * PT_PER_PX
        d = o_pt - w_pt
        print(f"  {i+1:>3} {w_pt:>8.2f} {o_pt:>8.2f} {d:>+7.2f}")
