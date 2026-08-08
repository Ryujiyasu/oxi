# -*- coding: utf-8 -*-
"""Read the area-chart probe: axis lines, gridlines, filled area polygons,
legend swatches and every text span."""
import sys, os

sys.stdout.reconfigure(encoding="utf-8")
import fitz

pdf = os.path.abspath(r"pipeline_data\pptx_probes\chart_area\chart_area.pdf")
doc = fitz.open(pdf)
NAMES = [
    "A1 area n=1 title", "A2 area n=2", "A3 area n=2 legend",
    "A4 stack n=2", "A5 stack n=2 legend", "A6 100% n=2 legend",
    "A7 area n=1 NOtitle", "A8 area n=3 legend",
]
print("frame sx=72 sy=72 sw=396 shh=288  -> right 468 bottom 360\n")
for pno in range(doc.page_count):
    page = doc[pno]
    print(f"===== S{pno+1}  {NAMES[pno]} =====")
    hl, vl, fills = [], [], []
    for d in page.get_drawings():
        for it in d["items"]:
            if it[0] == "l":
                a, b = it[1], it[2]
                if abs(a.y - b.y) < 0.2:
                    hl.append((round(a.y, 2), round(min(a.x, b.x), 2), round(max(a.x, b.x), 2)))
                elif abs(a.x - b.x) < 0.2:
                    vl.append((round(a.x, 2), round(min(a.y, b.y), 2), round(max(a.y, b.y), 2)))
        f = d.get("fill")
        if f:
            rgb = tuple(int(round(v * 255)) for v in f)
            pts = []
            for it in d["items"]:
                if it[0] == "l":
                    pts += [it[1], it[2]]
                elif it[0] == "re":
                    r = it[1]
                    pts += [fitz.Point(r.x0, r.y0), fitz.Point(r.x1, r.y1)]
            if pts:
                r = d["rect"]
                fills.append((rgb, round(r.x0, 2), round(r.y0, 2), round(r.x1, 2), round(r.y1, 2), len(pts), pts))
    hl = sorted(set(hl)); vl = sorted(set(vl))
    print("  H lines:", " ".join(f"y{y}[{x0}..{x1}]" for y, x0, x1 in hl[:14]))
    print("  V lines:", " ".join(f"x{x}[{y0}..{y1}]" for x, y0, y1 in vl[:14]))
    for rgb, x0, y0, x1, y1, n, pts in fills:
        big = (x1 - x0) > 40 and (y1 - y0) > 8
        tag = "AREA " if big else "swatch"
        print(f"  {tag} #{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}  x {x0:7.2f}..{x1:7.2f}  y {y0:7.2f}..{y1:7.2f}  pts={n}")
        if big:
            uniq = []
            for p in pts:
                t = (round(p.x, 2), round(p.y, 2))
                if not uniq or uniq[-1] != t:
                    uniq.append(t)
            print("        " + " ".join(f"({a},{b})" for a, b in uniq[:12]))
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                t = "".join(c["c"] for c in s["chars"])
                if t.strip():
                    print(f"  span ({s['origin'][0]:7.2f},{s['origin'][1]:7.2f}) x1={s['bbox'][2]:7.2f} sz={s['size']:5.2f} {s['font'][:20]:20s} {t!r}")
    print()
