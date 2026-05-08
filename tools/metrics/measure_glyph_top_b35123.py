"""Measure exact glyph top y in Word PNG vs Oxi PNG for b35123 row 1 cell.

Goal: confirm whether Word's text_y_offset uses sub-pixel precision
(= raw value, no round) or some other formula, and compare against Oxi's
post-Day-29 output.

Method: scan each PNG row by row, locate the first dark pixel cluster
in the cell content area, that's the glyph top y.
"""
from __future__ import annotations

import numpy as np
from PIL import Image
import sys


def find_glyph_top_in_region(png_path: str, dpi: int,
                              x_min_pt: float, x_max_pt: float,
                              y_min_pt: float, y_max_pt: float) -> float:
    """Find the topmost dark pixel within a rect, return as pt."""
    img = np.array(Image.open(png_path).convert('L'))
    x_min = int(x_min_pt * dpi / 72)
    x_max = int(x_max_pt * dpi / 72)
    y_min = int(y_min_pt * dpi / 72)
    y_max = int(y_max_pt * dpi / 72)
    region = img[y_min:y_max, x_min:x_max]
    dark = region < 200
    rows_with_dark = np.where(dark.any(axis=1))[0]
    if len(rows_with_dark) == 0:
        return None
    first_row = rows_with_dark[0]
    return (y_min + first_row) * 72.0 / dpi


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    DPI = 150

    # b35123 row 1 cell 0 region (Word DML: cell 0 at x=76.5pt; cell content
    # starts ~5.4pt right of cell border = x=82pt; lh region y=124-145pt).
    # Use a wide x window to catch any glyph in row 1.

    word_png = r'C:\Users\ryuji\oxi-main\pipeline_data\word_png\b35123fe8efc_tokumei_08_01\page_0001.png'
    oxi_pre  = r'C:\Users\ryuji\oxi-main\pipeline_data\oxi_png\b35123fe8efc_tokumei_08_01\page_0001.png'
    oxi_post = r'C:\tmp\b35123_d29b\page_p1.png'

    # Row 1: y region ~120-145, x ~80-130 (cell 0)
    print("=== Row 1 cell 0 glyph top y (x=80-130, y=120-148) ===")
    for label, p in [('Word PNG', word_png), ('Oxi pre-Day29', oxi_pre), ('Oxi post-Day29b', oxi_post)]:
        try:
            y = find_glyph_top_in_region(p, DPI, 80, 130, 120, 148)
            print(f"  {label:<22s}: glyph top y = {y}")
        except FileNotFoundError:
            print(f"  {label:<22s}: (file not found)")

    # Row 1 cell 1: x ~130-520
    print("\n=== Row 1 cell 1 glyph top y (x=140-520, y=120-148) ===")
    for label, p in [('Word PNG', word_png), ('Oxi pre-Day29', oxi_pre), ('Oxi post-Day29b', oxi_post)]:
        try:
            y = find_glyph_top_in_region(p, DPI, 140, 520, 120, 148)
            print(f"  {label:<22s}: glyph top y = {y}")
        except FileNotFoundError:
            print(f"  {label:<22s}: (file not found)")

    # Row 2 (Word DML y=144, Oxi y~146): x ~80-130
    print("\n=== Row 2 cell 0 glyph top y (x=80-130, y=140-170) ===")
    for label, p in [('Word PNG', word_png), ('Oxi pre-Day29', oxi_pre), ('Oxi post-Day29b', oxi_post)]:
        try:
            y = find_glyph_top_in_region(p, DPI, 80, 130, 140, 170)
            print(f"  {label:<22s}: glyph top y = {y}")
        except FileNotFoundError:
            print(f"  {label:<22s}: (file not found)")


if __name__ == "__main__":
    main()
