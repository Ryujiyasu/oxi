# -*- coding: utf-8 -*-
"""Read the Word line-chart PDF: chart frame + plot-area rects and the
series polyline (get_drawings items), plus axis/category/legend text
(rawdict spans)."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import fitz

base = r"pipeline_data\pptx_probes\chart_line"
pdf = os.path.join(base, "chart_line.pdf")

doc = fitz.open(pdf)
page = doc[0]
EMU = 12700.0  # EMU per pt

def pt_bounds(items):
    xs, ys = [], []
    for it in items:
        for seg in it[1:]:
            if hasattr(seg, "x"):
                xs.append(seg.x)
                ys.append(seg.y)
            elif isinstance(seg, (tuple, list)) and len(seg) == 2:
                xs.append(seg[0])
                ys.append(seg[1])
    return (min(xs), min(ys), max(xs), max(ys)) if xs else None


drawings = page.get_drawings()
print("=== drawings (%d) ===" % len(drawings))
for i, d in enumerate(drawings):
    r = d["rect"]
    fill = d.get("fill")
    stroke = d.get("color")
    sw = d.get("width")
    kinds = {}
    for it in d["items"]:
        kinds[it[0]] = kinds.get(it[0], 0) + 1
    b = pt_bounds(d["items"])
    print("[%d] rect=(%.2f,%.2f,%.2f,%.2f) fill=%s stroke=%s w=%s kinds=%s ptbounds=%s"
          % (i, r[0], r[1], r[2], r[3], fill, stroke, sw, kinds,
             tuple(round(v, 2) for v in b) if b else None))

print("=== text (rawdict spans) ===")
raw = page.get_text("rawdict")
for block in raw["blocks"]:
    if block["type"] != 0:
        continue
    for line in block["lines"]:
        for span in line["spans"]:
            txt = "".join(ch["c"] for ch in span["chars"])
            o = span["origin"]
            print("span font=%s size=%.2f origin=(%.2f,%.2f) text=%r"
                  % (span["font"], span["size"], o[0], o[1], txt))

print("page size:", page.rect.width, "x", page.rect.height)
