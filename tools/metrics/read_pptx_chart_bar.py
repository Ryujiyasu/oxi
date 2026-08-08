# -*- coding: utf-8 -*-
"""Measure Word's HORIZONTAL bar chart geometry from chart_bar.pdf.

Reads per-page:
  * every filled rect (bars) with its accent colour
  * every straight line item (axes, ticks, gridlines) from p["items"]
  * every text span (value-axis labels, category labels, auto title)
so the horizontal plot-area / bar-pitch / axis rules can be derived."""
import sys

sys.stdout.reconfigure(encoding="utf-8")
import fitz

PDF = r"pipeline_data\pptx_probes\chart_bar\chart_bar.pdf"


def hexc(c):
    if c is None:
        return None
    return "#%02X%02X%02X" % tuple(int(round(v * 255)) for v in c)


doc = fitz.open(PDF)
for pno in range(doc.page_count):
    page = doc[pno]
    print("=" * 72)
    print("PAGE", pno + 1, "size", page.rect)
    rects, hlines, vlines = [], [], []
    for p in page.get_drawings():
        for it in p["items"]:
            if it[0] == "re":
                r = it[1]
                rects.append((r, hexc(p.get("fill")), hexc(p.get("color"))))
            elif it[0] == "l":
                a, b = it[1], it[2]
                if abs(a.y - b.y) < 0.05:
                    hlines.append((a.y, min(a.x, b.x), max(a.x, b.x), hexc(p.get("color")), p.get("width")))
                elif abs(a.x - b.x) < 0.05:
                    vlines.append((a.x, min(a.y, b.y), max(a.y, b.y), hexc(p.get("color")), p.get("width")))
    print("-- filled rects (bars/frames) --")
    for r, f, c in rects:
        print(
            f"   x {r.x0:8.2f}..{r.x1:8.2f} (w {r.x1-r.x0:7.2f})  "
            f"y {r.y0:8.2f}..{r.y1:8.2f} (h {r.y1-r.y0:7.2f})  fill={f} stroke={c}"
        )
    print("-- horizontal lines --")
    for y, x0, x1, c, w in sorted(hlines):
        print(f"   y {y:8.2f}  x {x0:8.2f}..{x1:8.2f} (len {x1-x0:7.2f})  {c} w={w}")
    print("-- vertical lines --")
    for x, y0, y1, c, w in sorted(vlines):
        print(f"   x {x:8.2f}  y {y0:8.2f}..{y1:8.2f} (len {y1-y0:7.2f})  {c} w={w}")
    print("-- text spans --")
    d = page.get_text("rawdict")
    for blk in d["blocks"]:
        for ln in blk.get("lines", []):
            for sp in ln["spans"]:
                txt = "".join(c["c"] for c in sp.get("chars", []))
                if not txt.strip():
                    continue
                print(
                    f"   {txt!r:24s} font={sp['font']:22s} sz={sp['size']:6.2f} "
                    f"origin=({sp['origin'][0]:7.2f},{sp['origin'][1]:7.2f}) "
                    f"bbox=({sp['bbox'][0]:7.2f},{sp['bbox'][1]:7.2f},{sp['bbox'][2]:7.2f},{sp['bbox'][3]:7.2f})"
                )
