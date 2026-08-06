# -*- coding: utf-8 -*-
"""Dump individual bar rects from path items (fills can be merged into one bbox)."""
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
        fill = p.get("fill")
        for it in p["items"]:
            if it[0] == "re":  # rectangle item
                r = it[1]
                print(f"  bar rect=({r.x0:.2f},{r.y0:.2f},{r.x1:.2f},{r.y1:.2f}) "
                      f"w={r.width:.2f} h={r.height:.2f} fill={rgb01(fill)}")
    doc.close()
