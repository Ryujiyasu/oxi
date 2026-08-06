# -*- coding: utf-8 -*-
"""Dump ALL vector drawings + text from chart_legend.pdf — focus on the legend
swatches (small accent-coloured rects) and the legend labels (Revenue/Cost)."""
import fitz
doc = fitz.open(r"pipeline_data\pptx_probes\chart_legend\chart_legend.pdf")
page = doc[0]
paths = page.get_drawings()
print(f"n_paths={len(paths)}")
for i, p in enumerate(paths):
    fill = tuple(round(c, 3) for c in p["fill"]) if p.get("fill") is not None else None
    stroke = tuple(round(c, 3) for c in p["color"]) if p.get("color") is not None else None
    r = tuple(round(v, 2) for v in p["rect"])
    print(f"--- path[{i}] rect={r} fill={fill} stroke={stroke} items={len(p['items'])}")
    for it in p["items"]:
        print("   ", it)

# text (baseline origin + bbox + font + size) — text reconstructed from chars
print("\n--- text spans ---")
d = page.get_text("rawdict")
for b in d["blocks"]:
    for ln in b.get("lines", []):
        for sp in ln["spans"]:
            text = "".join(ch["c"] for ch in sp["chars"])
            print(f"   origin=({round(sp['origin'][0],2)},{round(sp['origin'][1],2)}) "
                  f"bbox=({round(sp['bbox'][0],2)},{round(sp['bbox'][1],2)},{round(sp['bbox'][2],2)},{round(sp['bbox'][3],2)}) "
                  f"font={sp['font']} size={round(sp['size'],2)} text={text!r}")
doc.close()
