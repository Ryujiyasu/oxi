"""Day 32 part 15 — Pixel-measure Oxi PNG glyph y to verify layout JSON y.

Day 32 part 14 found Word PDF glyph y vs Oxi layout JSON y differ by
+0.62pt for bd90b00 pi=2. But layout JSON y is the GDI TextOutW
reference point (= character cell top), which differs from PDF glyph
bbox top by internal leading (typically 1-3pt for fs=14).

This tool:
1. Loads Oxi-rendered PNG
2. For target paragraph, scans pixels in expected y/x range
3. Finds the topmost row with dark pixels (= actual glyph top in Oxi
   render)
4. Compares Oxi pixel-glyph-y vs Oxi layout JSON y → internal leading offset
5. With this correction, computes "true dy" between Oxi and Word PDF
   glyph y in same coordinate system

bd90b00 PNG: 1240×1754 at 150 DPI → 2.0833 pixels/pt.
"""
from __future__ import annotations
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')


def find_glyph_top(img_arr, x_range, y_range, threshold=180, min_dark_pixels=3):
    """Find topmost row in y_range with >= min_dark_pixels < threshold in x_range.

    Requires multiple consecutive dark pixels to filter anti-alias noise and
    avoid catching glyphs from adjacent paragraphs.
    """
    import numpy as np
    x_start, x_end = x_range
    y_start, y_end = y_range
    for y in range(y_start, y_end):
        row = img_arr[y, x_start:x_end]
        n_dark = int(np.sum(row < threshold))
        if n_dark >= min_dark_pixels:
            return y
    return None


def main():
    import numpy as np
    from PIL import Image

    png_path = r'pipeline_data/oxi_gdi_tmp/bd90b00ab7a7_order_05_p1.png'
    layout_path = r'C:\tmp\bd90b00ab7a7_order_05_v2_layout.json'

    print(f'Loading Oxi PNG: {png_path}')
    img = Image.open(png_path).convert('L')
    arr = np.array(img)
    print(f'  PNG size: {img.size}, dtype={arr.dtype}')

    # 1240 px × 8.27" → 150 DPI → 1pt = 2.0833 px
    scale = img.size[0] / (8.27 * 72.0)
    print(f'  Scale: {scale:.4f} pixels/pt')

    # Load Oxi layout JSON for ground-truth layout y values
    with open(layout_path, encoding='utf-8') as f:
        layout = json.load(f)

    # Find page 1 first text element per para_idx (matches bd90b00 pi=1, 2, 4)
    page1 = layout['pages'][0]
    by_pi = {}
    for el in page1.get('elements', []):
        if el.get('type') != 'text':
            continue
        pi = el.get('para_idx')
        y = el.get('y', 0)
        x = el.get('x', 0)
        text = el.get('text', '')
        if pi not in by_pi or y < by_pi[pi]['y']:
            by_pi[pi] = {'y': y, 'x': x, 'text': text}

    # Word data (from Day 32 part 13 PDF measurement)
    word_data = {
        # word_i: (anchor_y, glyph_top_y_pdf)
        2: (75.00, 76.88),    # （統計法...
        4: (117.00, 119.12),  # 厚生労働大臣 殿
        9: (197.00, 198.55),  # 連絡先e-mail
    }

    # Map Word i → Oxi pi (Day 32 part 5 finding: Oxi pi = Word i - 1 only outside table)
    # Verified bd90b00: Word i=2 → Oxi pi=1, Word i=4 → Oxi pi=3, Word i=9 → Oxi pi=8
    # Actually from part 14 data, bd90b00 layout: pi=1 y=77.5 = Word i=2 mapping
    targets = [
        # (label, oxi_pi_estimate, word_i, fs)
        ('（統計法...', 1, 2, 14.0),
        ('厚生労働大臣', 3, 4, 14.0),
        ('連絡先e-mail', 8, 9, 10.5),
    ]

    print(f'\n{"label":<15} {"oxi_pi":>6} {"oxi_y":>7} {"px_scan_y":>10} {"px_glyph_y_pt":>14} {"oxi_il":>7} {"word_anchor":>12} {"word_pdf":>9} {"oxi-word":>9}')
    for label, oxi_pi, word_i, fs in targets:
        if oxi_pi not in by_pi:
            print(f'  {label}: oxi_pi={oxi_pi} not found in layout')
            continue
        oxi_y = by_pi[oxi_pi]['y']
        oxi_x = by_pi[oxi_pi]['x']
        # Scan pixel range narrow around expected y (no overlap with prev paragraph)
        y_search_start = int((oxi_y - 1.5) * scale)
        y_search_end = int((oxi_y + 8) * scale)
        # X range: just first 1 character (~ fs in pt)
        x_start_px = max(0, int(oxi_x * scale))
        x_end_px = min(arr.shape[1], int((oxi_x + fs) * scale))
        glyph_top_px = find_glyph_top(arr, (x_start_px, x_end_px),
                                      (y_search_start, y_search_end))
        if glyph_top_px is None:
            print(f'  {label}: no dark pixel found in scan range')
            continue
        glyph_top_pt = glyph_top_px / scale
        oxi_internal_leading = round(glyph_top_pt - oxi_y, 2)
        wd_anchor, wd_pdf = word_data.get(word_i, (None, None))
        if wd_pdf is None:
            print(f'  {label}: no Word PDF data for word_i={word_i}')
            continue
        # Compare Oxi actual glyph y vs Word PDF glyph y (same coordinate system)
        true_dy = round(glyph_top_pt - wd_pdf, 2)
        print(f'  {label:<15} {oxi_pi:>6} {oxi_y:>7.2f} {glyph_top_px:>10} {glyph_top_pt:>14.2f} {oxi_internal_leading:>+7.2f} {wd_anchor:>12.2f} {wd_pdf:>9.2f} {true_dy:>+9.2f}')

    print('\nInterpretation:')
    print('  oxi_il = pixel glyph top - layout JSON y (= GDI internal leading)')
    print('  oxi-word = oxi pixel glyph top - Word PDF glyph top (in same coord system)')
    print('  Small oxi-word (< 1pt) → Oxi/Word match in actual rendering')
    print('  Large oxi-word → real bug between Oxi and Word')


if __name__ == '__main__':
    main()
