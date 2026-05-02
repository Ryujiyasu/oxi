"""Pixel-precise Y diff between Word and Oxi PNG for 2ea81a page 1.

Strategy:
  1. Find horizontal black/dark lines in each image (table borders)
  2. Match top-of-first-table by finding the first significant horizontal
     line with > N dark pixels
  3. Compute Y delta in pixels, convert to pt (150dpi → 1pt = ~2.083px)
"""
from PIL import Image
import numpy as np

WORD_PNG = "pipeline_data/word_png/2ea81a8441cc_0025006-192/page_0001.png"
OXI_PNG = "pipeline_data/oxi_png/2ea81a8441cc_0025006-192/page_0001.png"

# 150 DPI, 1pt = 1/72 inch = 150/72 pixels ≈ 2.083 px/pt
DPI = 150
PT_PER_PX = 72.0 / DPI


def find_horizontal_lines(img_path, min_dark_count=200):
    """Return list of (y_px, dark_count) for rows with significant dark pixels."""
    img = Image.open(img_path).convert("L")
    arr = np.array(img)
    # dark pixel = brightness < 128
    dark_mask = arr < 128
    row_dark_counts = dark_mask.sum(axis=1)
    h, w = arr.shape
    # Find rows with > min_dark_count dark pixels
    lines = []
    in_run = False
    run_start = 0
    for y in range(h):
        if row_dark_counts[y] > min_dark_count:
            if not in_run:
                in_run = True
                run_start = y
        else:
            if in_run:
                in_run = False
                # use middle of run as line position
                mid = (run_start + y) // 2
                lines.append((mid, max(row_dark_counts[run_start:y])))
    return lines, h, w


def main():
    print(f"Resolution: {DPI} DPI, {PT_PER_PX:.4f} pt/pixel")
    print()
    for label, path in [("Word", WORD_PNG), ("Oxi", OXI_PNG)]:
        lines, h, w = find_horizontal_lines(path, min_dark_count=300)
        print(f"{label} ({path}):")
        print(f"  size: {w}×{h}px, {h * PT_PER_PX:.1f}pt high")
        print(f"  Found {len(lines)} significant horizontal lines (top 10):")
        for y_px, dark in lines[:10]:
            y_pt = y_px * PT_PER_PX
            print(f"    y_px={y_px} (y_pt={y_pt:.1f}pt) dark_count={dark}")


if __name__ == "__main__":
    main()
