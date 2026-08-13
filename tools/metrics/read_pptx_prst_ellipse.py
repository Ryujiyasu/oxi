# -*- coding: utf-8 -*-
"""Read PowerPoint's exported geometry for the preset-geometry probe.

An `a:prstGeom` ellipse is emitted by PowerPoint as a closed path of four cubic
Beziers.  What we need for the renderer is:
  * does the path's extent equal the shape box exactly (fill), and where does a
    stroke sit relative to it (inside / centred / outside)?
  * is the Bezier control offset the standard circle constant (kappa = 0.5523)?
  * how does rotation / flip transform the path?

Bezier control points bulge OUTSIDE the true curve, so `p["rect"]` overstates
the extent -- take the extent from the on-curve endpoints only.
"""
import json
import os
import sys

import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DIR = os.path.abspath(r"pipeline_data\pptx_probes\prst_ellipse")
PDF = os.path.join(DIR, "prst_ellipse.pdf")
NAMES = [
    "E1_ellipse_fill", "E2_ellipse_line", "E3_circle", "E4_ellipse_flat",
    "E5_ellipse_rot30", "E6_ellipse_fliph", "E7_ellipse_text", "E8_roundrect",
    "E9_homeplate", "E10_teardrop",
]
BOX = {  # what the deck declared (pt)
    "E1_ellipse_fill": (72, 72, 396, 288), "E2_ellipse_line": (72, 72, 396, 288),
    "E3_circle": (72, 72, 288, 288), "E4_ellipse_flat": (72, 72, 396, 108),
    "E5_ellipse_rot30": (72, 72, 396, 288), "E6_ellipse_fliph": (72, 72, 396, 288),
    "E7_ellipse_text": (72, 72, 396, 288), "E8_roundrect": (72, 72, 396, 288),
    "E9_homeplate": (72, 72, 396, 288), "E10_teardrop": (72, 72, 396, 288),
}


def oncurve(items):
    """Extent from on-curve points only (Bezier controls bulge outside)."""
    xs, ys = [], []
    for it in items:
        k = it[0]
        if k == "l":
            for p in (it[1], it[2]):
                xs.append(p.x); ys.append(p.y)
        elif k == "c":
            for p in (it[1], it[4]):   # start and end are on-curve
                xs.append(p.x); ys.append(p.y)
        elif k == "re":
            r = it[1]
            xs += [r.x0, r.x1]; ys += [r.y0, r.y1]
    if not xs:
        return None
    return (min(xs), min(ys), max(xs), max(ys))


def main():
    doc = fitz.open(PDF)
    for i, page in enumerate(doc):
        nm = NAMES[i] if i < len(NAMES) else "slide%d" % (i + 1)
        bx, by, bw, bh = BOX.get(nm, (0, 0, 0, 0))
        print("\n=== %s   box=(%g,%g,%g,%g) ===" % (nm, bx, by, bw, bh))
        for d in page.get_drawings():
            items = d.get("items") or []
            kinds = {}
            for it in items:
                kinds[it[0]] = kinds.get(it[0], 0) + 1
            f = d.get("fill")
            s = d.get("color")
            if f is None and s is None:
                continue
            e = oncurve(items)
            if e is None:
                continue
            # skip the page-sized white background
            if e[2] - e[0] > 700:
                continue
            fs = "-" if f is None else "#%02X%02X%02X" % tuple(int(round(c * 255)) for c in f)
            ss = "-" if s is None else "#%02X%02X%02X" % tuple(int(round(c * 255)) for c in s)
            print("   items=%s fill=%s stroke=%s w=%s" % (kinds, fs, ss, d.get("width")))
            print("      on-curve extent  x %.2f..%.2f (w %.2f)   y %.2f..%.2f (h %.2f)"
                  % (e[0], e[2], e[2] - e[0], e[1], e[3], e[3] - e[1]))
            print("      vs box           dx0 %+.2f dy0 %+.2f dx1 %+.2f dy1 %+.2f"
                  % (e[0] - bx, e[1] - by, e[2] - (bx + bw), e[3] - (by + bh)))
            if kinds.get("c") == 4 and len(items) == 4:
                # kappa from the first Bezier: control offset / semi-axis
                it = items[0]
                p0, c1, c2, p3 = it[1], it[2], it[3], it[4]
                for lbl, a, b, semi in (
                    ("c1", c1, p0, (e[2] - e[0]) / 2.0),
                    ("c2", c2, p3, (e[3] - e[1]) / 2.0),
                ):
                    dx, dy = abs(a.x - b.x), abs(a.y - b.y)
                    off = max(dx, dy)
                    if semi:
                        print("      kappa(%s) = %.4f" % (lbl, off / semi))
    doc.close()


if __name__ == "__main__":
    main()
