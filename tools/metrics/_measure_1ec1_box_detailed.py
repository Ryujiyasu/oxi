# -*- coding: utf-8 -*-
"""Detailed PDF-level measurement of □ glyph positions in 1ec1.

For each □ paragraph: compare PyMuPDF search bbox.x0 vs leftmost dark pixel
to understand glyph bearing and confirm where Word actually places content."""
import sys, os
import fitz

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
PDF = "pipeline_data/1ec1_actual_pdf_measure/1ec1.pdf"
BOX = '□'

d = fitz.open(PDF)
print(f"Pages: {d.page_count}")

# Page 1 has page margin top=1134tw=56.7pt, left=851tw=42.55pt
LEFT_MARGIN_PT = 851 / 20  # = 42.55pt
print(f"Expected page left margin: {LEFT_MARGIN_PT}pt")

zoom = 4.0
for page_idx in range(min(d.page_count, 3)):
    page = d[page_idx]
    instances = page.search_for(BOX)
    if not instances:
        continue
    print(f"\n===== Page {page_idx+1}: {len(instances)} □ instances =====")
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    w, h, n = pix.width, pix.height, pix.n
    s = pix.samples

    for i, inst in enumerate(instances[:6]):
        # Search bbox
        bx0, by0, bx1, by1 = inst.x0, inst.y0, inst.x1, inst.y1
        # Find leftmost dark pixel within rect's vertical range,
        # scanning from page.left=0 to a bit past glyph's right
        top_px = int(by0 * zoom)
        bottom_px = int(by1 * zoom)
        # Scan only within a narrow window around the bbox (avoid catching unrelated rules)
        left_search = max(0, int((bx0 - 2) * zoom))
        right_search = min(w, int((bx1 + 1) * zoom))
        leftmost = None
        for py in range(max(0, top_px), min(h, bottom_px)):
            for px in range(left_search, right_search):
                off = (py * w + px) * n
                r, g, bb = s[off], s[off+1], s[off+2]
                if r < 200 and g < 200 and bb < 200:
                    if leftmost is None or px < leftmost:
                        leftmost = px
                    break
        leftmost_pt = leftmost / zoom if leftmost else None
        # Bearing = visible - search_x0
        bearing = leftmost_pt - bx0 if leftmost_pt else None
        # Excess = search_x0 - margin
        excess = bx0 - LEFT_MARGIN_PT
        print(f"  □#{i+1}: search bbox L={bx0:.2f} R={bx1:.2f} | dark pixel x={leftmost} ({leftmost_pt:.2f}pt) | bearing={bearing:.2f}pt | excess_over_margin={excess:.2f}pt")

d.close()
