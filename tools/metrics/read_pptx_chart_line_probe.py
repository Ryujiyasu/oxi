# -*- coding: utf-8 -*-
"""Measure each page of the line-chart probe PDF: summarize the drawings that
define the plot area (Y axis / X axis / gridlines) per page.

Prints per page: axes + gridline bounds + series polyline bounds + markers +
text spans (value labels / title / legend)."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import fitz

pdf = r"pipeline_data\pptx_probes\chart_line_probe\chart_line_probe.pdf"

def pt(v):
    return (v.x, v.y) if hasattr(v, "x") else v

def pt_bounds(items):
    xs, ys = [], []
    for it in items:
        k = it[0]
        for q in it[1:]:
            if isinstance(q, fitz.Point):
                xs.append(q.x); ys.append(q.y)
            elif isinstance(q, (tuple, list)) and len(q) == 2:
                xs.append(q[0]); ys.append(q[1])
            elif isinstance(q, fitz.Rect):
                xs += [q.x0, q.x1]; ys += [q.y0, q.y1]
    return (min(xs), min(ys), max(xs), max(ys)) if xs else None

doc = fitz.open(pdf)
for pno, page in enumerate(doc):
    print("=" * 40, "PAGE", pno)
    dw = page.get_drawings()
    for i, d in enumerate(dw):
        r = d["rect"]
        w = d["width"]
        print(f"[{i}] rect=({r.x0:.2f},{r.y0:.2f},{r.x1:.2f},{r.y1:.2f}) "
              f"fill={d['fill']} stroke={d['color']} w={'-' if w is None else round(w,2)} "
              f"kinds={ {k: c for k, c in __import__('collections').Counter(it[0] for it in d['items']).items()} } "
              f"ptb={None if pt_bounds(d['items']) is None else tuple(round(v,2) for v in pt_bounds(d['items']))}")
    spans = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                txt = "".join(ch["c"] for ch in s.get("chars", []))
                spans.append((s["font"], s["size"], tuple(round(v,2) for v in s["origin"]), txt))
    print("  spans:")
    for s in spans:
        print("   ", s)
