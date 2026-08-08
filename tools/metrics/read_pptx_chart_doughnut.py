# -*- coding: utf-8 -*-
"""Export chart_doughnut via PowerPoint COM and read Word's ring geometry.

A doughnut slice is an annular sector: its path carries points on the OUTER
arc and on the INNER arc.  The union bbox of all slices is the outer circle,
which gives the centre and outer radius; the minimum point distance gives the
hole radius; the arc endpoints give the slice angles (0 deg = 12 o'clock,
clockwise, matching the pie derivation)."""
import sys, os, math

sys.stdout.reconfigure(encoding="utf-8")
import win32com.client
import fitz

BASE = r"pipeline_data\pptx_probes\chart_doughnut"
PPTX = os.path.abspath(os.path.join(BASE, "chart_doughnut.pptx"))
PDF = os.path.abspath(os.path.join(BASE, "chart_doughnut.pdf"))

if not os.path.exists(PDF) or os.path.getmtime(PDF) < os.path.getmtime(PPTX):
    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = ppt.Presentations.Open(PPTX, WithWindow=False)
        pres.SaveAs(PDF, 32)
        pres.Close()
    finally:
        ppt.Quit()
    print("exported:", PDF)

SX, SY, SW, SH = 72.0, 72.0, 396.0, 288.0
ACCENT = {(0.31, 0.506, 0.741): 1, (0.753, 0.314, 0.302): 2, (0.608, 0.733, 0.349): 3}


def ang(cx, cy, x, y):
    """Degrees clockwise from 12 o'clock."""
    return (math.degrees(math.atan2(x - cx, cy - y)) + 360.0) % 360.0


doc = fitz.open(PDF)
for pno in range(doc.page_count):
    page = doc[pno]
    print(f"\n===== slide {pno + 1} =====")

    slices, swatches = [], []
    for p in page.get_drawings():
        f = p.get("fill")
        key = tuple(round(c, 3) for c in f) if f else None
        if key not in ACCENT:
            continue
        pts = []
        for it in p["items"]:
            if it[0] == "c":
                pts.extend([it[1], it[4]])
            elif it[0] == "l":
                pts.extend([it[1], it[2]])
            elif it[0] == "re":
                r = it[1]
                if r.width < 12 and r.height < 12:
                    swatches.append((round(r.x0, 2), round(r.y0, 2), ACCENT[key]))
                pts = []
                break
        if pts:
            slices.append((ACCENT[key], p["rect"], pts))

    if slices:
        x0 = min(s[1].x0 for s in slices); x1 = max(s[1].x1 for s in slices)
        y0 = min(s[1].y0 for s in slices); y1 = max(s[1].y1 for s in slices)
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        r_out = ((x1 - x0) + (y1 - y0)) / 4
        print(f"  ring bbox x[{x0:.2f},{x1:.2f}] y[{y0:.2f},{y1:.2f}]")
        print(f"  centre=({cx:.2f},{cy:.2f}) (= sx+{cx-SX:.2f}, sy+{cy-SY:.2f})  r_out={r_out:.2f}")
        print(f"  top={y0:.2f} (= sy+{y0-SY:.2f})  bot={y1:.2f} (= sy+shh-{SY+SH-y1:.2f})")
        rr = []
        for idx, rect, pts in sorted(slices):
            d = [math.hypot(p.x - cx, p.y - cy) for p in pts]
            a = [ang(cx, cy, p.x, p.y) for p in pts]
            rr.append(min(d))
            print(f"  accent{idx}: r_in={min(d):6.2f} r_out={max(d):6.2f} "
                  f"ang[{min(a):6.2f},{max(a):6.2f}] n={len(pts)}")
        print(f"  hole ratio r_in/r_out = {sum(rr)/len(rr)/r_out:.4f}")

    for s in sorted(swatches):
        print(f"  legend swatch x0={s[0]:7.2f} y0={s[1]:7.2f} accent{s[2]}")

    for blk in page.get_text("rawdict")["blocks"]:
        for ln in blk.get("lines", []):
            for sp in ln["spans"]:
                t = "".join(c["c"] for c in sp.get("chars", []))
                if t.strip():
                    print(f"  span y={sp['origin'][1]:7.2f} x0={sp['origin'][0]:7.2f} "
                          f"x1={sp['bbox'][2]:7.2f} sz={sp['size']:5.2f} "
                          f"{sp['font']:<24s} {t.strip()!r}")
