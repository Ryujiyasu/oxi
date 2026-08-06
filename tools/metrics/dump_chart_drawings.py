# -*- coding: utf-8 -*-
"""Dump chart drawing structure (rects with fill colors + line items) for a PDF page."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
import fitz


def rgb01(c):
    if not c:
        return None
    return tuple(round(x, 3) for x in c)


for path in sys.argv[1:]:
    print(f"=== {path}")
    doc = fitz.open(path)
    page = doc[0]
    for p in page.get_drawings():
        r = p["rect"]
        # only report filled rects (bars) and horizontal/vertical lines of the plot
        fill = p.get("fill")
        print(f"  rect=({r.x0:.2f},{r.y0:.2f},{r.x1:.2f},{r.y1:.2f}) "
              f"w={r.width:.2f} h={r.height:.2f} fill={rgb01(fill)} "
              f"type={p.get('type')}")
        # count line items separately
        items = p["items"]
        if len(items) and all(it[0] == "l" for it in items):
            # it's a polyline: report first/last points
            pts = [it[1] for it in items]
            xs = [pt[0] for pt in pts]
            ys = [pt[1] for pt in pts]
            print(f"    line n={len(pts)} x:[{min(xs):.2f},{max(xs):.2f}] "
                  f"y:[{min(ys):.2f},{max(ys):.2f}]")
    doc.close()
