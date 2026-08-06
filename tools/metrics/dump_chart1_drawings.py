# -*- coding: utf-8 -*-
"""Dump ALL vector drawings from chart1.pdf (check for legend swatch near
'Series 1' at y~100)."""
import fitz
doc = fitz.open(r"pipeline_data\pptx_probes\chart1\chart1.pdf")
paths = doc[0].get_drawings()
print(f"n_paths={len(paths)}")
for i, p in enumerate(paths):
    fill = tuple(round(c, 3) for c in p["fill"]) if p.get("fill") is not None else None
    stroke = tuple(round(c, 3) for c in p["color"]) if p.get("color") is not None else None
    print(f"--- path[{i}] fill={fill} stroke={stroke} items={len(p['items'])}")
    for it in p["items"]:
        print("   ", it)
doc.close()
