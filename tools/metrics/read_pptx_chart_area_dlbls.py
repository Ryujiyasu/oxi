# -*- coding: utf-8 -*-
"""Read the AREA data-label probe: plot geometry (axis lines), the area
fill polygons' vertices, every text span (value axis labels, category
labels, data labels) and the legend swatches.

Data labels are the only spans that sit INSIDE the plot rectangle, so they
are reported separately with their baseline/centre so the placement rule
(centred on the point? offset above? inside the band?) can be read off."""
import sys, os

sys.stdout.reconfigure(encoding="utf-8")
import fitz

pdf = os.path.abspath(r"pipeline_data\pptx_probes\chart_area_dlbls\chart_area_dlbls.pdf")
doc = fitz.open(pdf)
NAMES = [
    "D1 area n=1", "D2 area n=2", "D3 stack n=2", "D4 100% n=2",
    "D5 area n=1 0.0%", "D6 area n=2 legend band", "D7 area n=2 legend bare",
]
SX, SY, SW, SHH = 72.0, 72.0, 396.0, 288.0
print(f"frame sx={SX} sy={SY} sw={SW} shh={SHH} -> right {SX+SW} bottom {SY+SHH}\n")

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
            r = d["rect"]
            pts = []
            for it in d["items"]:
                if it[0] == "l":
                    pts += [it[1], it[2]]
            fills.append((rgb, r, pts))
    hl = sorted(set(hl)); vl = sorted(set(vl))
    print("  H lines:", " ".join(f"y{y}[{x0}..{x1}]" for y, x0, x1 in hl[:12]))
    print("  V lines:", " ".join(f"x{x}[{y0}..{y1}]" for x, y0, y1 in vl[:12]))

    plot_left = min((x for x, _, _ in vl), default=None)
    plot_bot = max((y for y, _, _ in hl), default=None)
    plot_top = min((y for y, _, _ in hl), default=None)
    plot_right = max((x1 for _, _, x1 in hl), default=None)
    print(f"  plot: left {plot_left} top {plot_top} right {plot_right} bot {plot_bot}"
          f"  (sx+{None if plot_left is None else round(plot_left-SX,2)},"
          f" sy+{None if plot_top is None else round(plot_top-SY,2)})")

    for rgb, r, pts in fills:
        big = (r.x1 - r.x0) > 40 and (r.y1 - r.y0) > 8
        tag = "AREA  " if big else "swatch"
        print(f"  {tag} #{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
              f"  x {r.x0:7.2f}..{r.x1:7.2f}  y {r.y0:7.2f}..{r.y1:7.2f}")
        if big and pts:
            seen, uniq = set(), []
            for p in pts:
                k = (round(p.x, 2), round(p.y, 2))
                if k not in seen:
                    seen.add(k); uniq.append(k)
            uniq.sort()
            print("        verts:", " ".join(f"({x:.2f},{y:.2f})" for x, y in uniq[:12]))

    inside, outside = [], []
    for blk in page.get_text("rawdict")["blocks"]:
        for ln in blk.get("lines", []):
            for sp in ln["spans"]:
                t = "".join(c["c"] for c in sp.get("chars", [])).strip()
                if not t:
                    continue
                ox, oy = sp["origin"]
                bb = sp["bbox"]
                rec = (t, round(ox, 2), round(oy, 2), round(bb[0], 2), round(bb[2], 2),
                       round(sp["size"], 2), sp["font"])
                if (plot_left is not None and plot_left - 2 <= ox <= plot_right + 2
                        and plot_top - 2 <= oy <= plot_bot + 2):
                    inside.append(rec)
                else:
                    outside.append(rec)
    print("  -- spans INSIDE plot (data labels) --")
    for t, ox, oy, x0, x1, fs, fnt in sorted(inside, key=lambda r: (r[2], r[1])):
        print(f"     '{t}'  origin=({ox},{oy})  x {x0}..{x1} (w {round(x1-x0,2)})  fs {fs}  {fnt}")
    print("  -- spans OUTSIDE plot --")
    for t, ox, oy, x0, x1, fs, fnt in sorted(outside, key=lambda r: (r[2], r[1])):
        print(f"     '{t}'  origin=({ox},{oy})  x {x0}..{x1}  fs {fs}  {fnt}")
    print()
