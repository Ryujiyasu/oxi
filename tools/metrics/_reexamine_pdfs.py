# -*- coding: utf-8 -*-
"""Re-examine all bisection PDFs without the [:4] limit bug."""
import sys, os, glob, fitz
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DIRS = [
    r'C:\Users\ryuji\oxi-4\pipeline_data\1ec1_body_bisect',
    r'C:\Users\ryuji\oxi-4\pipeline_data\1ec1_body_bisect_after',
]

for d in DIRS:
    for pdf in sorted(glob.glob(os.path.join(d, '*.pdf'))):
        name = os.path.splitext(os.path.basename(pdf))[0]
        try:
            doc = fitz.open(pdf)
            boxes = []
            for pi in range(doc.page_count):
                for inst in doc[pi].search_for('□'):
                    boxes.append((pi+1, inst.x0, inst.y0))
            doc.close()
        except: continue
        if not boxes: continue
        boxes_sorted = sorted(boxes, key=lambda b: b[2])
        # Group by x-bucket: x≈46.1 (Shape 35 left=0), x≈46.6 (Shape 35 BOX[3]/[4] OR Shape 9 P1/P3 with left=105 in NO-trigger context), x≈55.3 (Shape 9 P1 in trigger context OR Shape 9 P6 firstLine)
        # Identify Shape 9 P1: it's the first box AFTER Shape 35 boxes (x in 46.0-46.7)
        # Heuristic: find boxes with x > 46.3 (excludes Shape 35 BOX[1]/[2] at 46.08), then first one is Shape 9 P1 OR Shape 35 BOX[3]/[4]
        # Better: the NUMBER of distinct x clusters tells us. 2 clusters → Shape 9 alone or with Shape 35; 3 clusters → trigger active
        x_set = set(round(b[1], 1) for b in boxes)
        # Print all boxes
        print(f"\n{name} ({len(boxes)} □ in {len(x_set)} x clusters):")
        for b in boxes_sorted:
            print(f"   x={b[1]:.2f}  y={b[2]:.2f}  P{b[0]}")
