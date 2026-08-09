# -*- coding: utf-8 -*-
"""Read the SCATTER probe's Word render-truth.

Scatter has a NUMERIC x axis, so the readout reports the axis lines, the
tick label spans on both axes, every marker (small fill) with its centre,
the polylines (S3/S4/S5) and any spans sitting inside the plot rect
(data labels)."""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
import fitz

PDF = os.path.abspath(r"pipeline_data\pptx_probes\chart_scatter\chart_scatter.pdf")
NAMES = [
    "S1 markers 1ser", "S2 markers 2ser", "S3 lines+mk 1ser",
    "S4 smooth+mk 1ser", "S5 lines only 2ser", "S6 markers 2ser legend",
    "S7 markers 1ser dlbls", "S8 markers 1ser notitle",
]
SX, SY, SW, SHH = 72.0, 72.0, 396.0, 288.0


def main() -> None:
    doc = fitz.open(PDF)
    print(f"frame sx={SX} sy={SY} sw={SW} shh={SHH} "
          f"-> right {SX+SW} bottom {SY+SHH}\n")
    for pno in range(doc.page_count):
        page = doc[pno]
        print(f"===== S{pno+1}  {NAMES[pno]} =====")
        hl, vl, marks, polys = [], [], [], []
        for d in page.get_drawings():
            f = d.get("fill")
            r = d["rect"]
            pts = []
            for it in d["items"]:
                if it[0] == "l":
                    a, b = it[1], it[2]
                    pts += [a, b]
                    if abs(a.y - b.y) < 0.2:
                        hl.append((round(a.y, 2), round(min(a.x, b.x), 2),
                                   round(max(a.x, b.x), 2)))
                    elif abs(a.x - b.x) < 0.2:
                        vl.append((round(a.x, 2), round(min(a.y, b.y), 2),
                                   round(max(a.y, b.y), 2)))
            w, h = r.x1 - r.x0, r.y1 - r.y0
            if f and 2 < w < 20 and 2 < h < 20:
                rgb = tuple(int(round(v * 255)) for v in f)
                marks.append((rgb, round((r.x0 + r.x1) / 2, 2),
                              round((r.y0 + r.y1) / 2, 2),
                              round(w, 2), round(h, 2), len(d["items"])))
            elif not f and len(pts) >= 4 and (w > 30 or h > 30):
                col = d.get("color")
                rgb = tuple(int(round(v * 255)) for v in col) if col else None
                seen, uniq = set(), []
                for p in pts:
                    k = (round(p.x, 2), round(p.y, 2))
                    if k not in seen:
                        seen.add(k)
                        uniq.append(k)
                polys.append((rgb, round(d.get("width") or 0, 2), uniq))

        hl = sorted(set(hl))
        vl = sorted(set(vl))
        plot_left = min((x for x, _, _ in vl), default=None)
        plot_bot = max((y for y, _, _ in hl), default=None)
        plot_top = min((y for y, _, _ in hl), default=None)
        plot_right = max((x1 for _, _, x1 in hl), default=None)
        print(f"  plot: left {plot_left} top {plot_top} right {plot_right} "
              f"bot {plot_bot}  (sx+{None if plot_left is None else round(plot_left-SX,2)},"
              f" sy+{None if plot_top is None else round(plot_top-SY,2)})")
        print("  H lines:", " ".join(f"y{y}[{x0}..{x1}]" for y, x0, x1 in hl[:14]))
        print("  V lines:", " ".join(f"x{x}[{y0}..{y1}]" for x, y0, y1 in vl[:14]))

        for rgb, cx, cy, w, h, n in sorted(marks, key=lambda m: (m[1], m[2])):
            print(f"  MARK #{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X} "
                  f"c=({cx},{cy}) {w}x{h} items={n}")
        for rgb, wd, uniq in polys:
            tag = f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}" if rgb else "none"
            print(f"  POLY {tag} w={wd} pts=" +
                  " ".join(f"({x:.2f},{y:.2f})" for x, y in uniq[:10]))

        inside, outside = [], []
        for blk in page.get_text("rawdict")["blocks"]:
            for ln in blk.get("lines", []):
                for sp in ln["spans"]:
                    t = "".join(c["c"] for c in sp.get("chars", [])).strip()
                    if not t:
                        continue
                    ox, oy = sp["origin"]
                    bb = sp["bbox"]
                    rec = (t, round(ox, 2), round(oy, 2), round(bb[0], 2),
                           round(bb[2], 2), round(sp["size"], 2), sp["font"])
                    if (plot_left is not None
                            and plot_left - 2 <= ox <= plot_right + 2
                            and plot_top - 2 <= oy <= plot_bot + 2):
                        inside.append(rec)
                    else:
                        outside.append(rec)
        print("  -- spans INSIDE plot --")
        for t, ox, oy, x0, x1, fs, fnt in sorted(inside, key=lambda r: (r[2], r[1])):
            print(f"     '{t}'  origin=({ox},{oy})  x {x0}..{x1} "
                  f"(w {round(x1-x0,2)})  fs {fs}  {fnt}")
        print("  -- spans OUTSIDE plot --")
        for t, ox, oy, x0, x1, fs, fnt in sorted(outside, key=lambda r: (r[2], r[1])):
            print(f"     '{t}'  origin=({ox},{oy})  x {x0}..{x1}  fs {fs}  {fnt}")
        print()


if __name__ == "__main__":
    main()
