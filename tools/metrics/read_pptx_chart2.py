# -*- coding: utf-8 -*-
"""Chart spec wave-2 read: measure chart2.pdf with fitz — text spans + vector
drawings (bars/axes), printing per-item so the multi-series cluster layout,
per-series colours, and legend positions are fully exposed."""
import sys, os, json
sys.stdout.reconfigure(encoding="utf-8")
import fitz

base = r"pipeline_data\pptx_probes\chart2"
pdf_path = os.path.join(base, "chart2.pdf")

doc = fitz.open(pdf_path)
page = doc[0]
print("page rect:", page.rect)

print("\n=== TEXT SPANS ===")
d = page.get_text("rawdict")
for block in d["blocks"]:
    if block["type"] != 0:
        continue
    for line in block["lines"]:
        for span in line["spans"]:
            txt = "".join(c["c"] for c in span["chars"])
            if not txt.strip():
                continue
            o = span["origin"]
            col = span["color"]
            size = span["size"]
            font = span["font"]
            print(f"  '{txt}' x={o[0]:.2f} y={o[1]:.2f} size={size:.2f} color=#{col:06x} font={font}")

print("\n=== VECTOR DRAWINGS (per item) ===")
paths = page.get_drawings()
print("n_paths:", len(paths))
for i, p in enumerate(paths):
    r = p["rect"]
    fill = p.get("fill")
    stroke = p.get("color")
    print(f"  path[{i}] rect=({r.x0:.1f},{r.y0:.1f},{r.x1:.1f},{r.y1:.1f}) "
          f"w={r.width:.1f} h={r.height:.1f} fill={fill} stroke={stroke} "
          f"n_items={len(p['items'])}")
    for it in p["items"]:
        print(f"      item {it}")
