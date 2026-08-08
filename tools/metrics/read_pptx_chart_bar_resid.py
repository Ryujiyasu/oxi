# -*- coding: utf-8 -*-
"""Export chart_bar_resid via PowerPoint COM and read Word's render-truth for
the two unmeasured horizontal-bar values: plot_top under an explicit title,
and stacked data-label placement."""
import sys, os

sys.stdout.reconfigure(encoding="utf-8")
import win32com.client
import fitz

BASE = r"pipeline_data\pptx_probes\chart_bar_resid"
PPTX = os.path.abspath(os.path.join(BASE, "chart_bar_resid.pptx"))
PDF = os.path.abspath(os.path.join(BASE, "chart_bar_resid.pdf"))

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
doc = fitz.open(PDF)
for pno in range(doc.page_count):
    page = doc[pno]
    print(f"\n===== slide {pno + 1} =====")

    fills, hticks, vticks = [], [], []
    for p in page.get_drawings():
        for it in p["items"]:
            if it[0] == "re":
                r = it[1]
                if p.get("fill") is not None and r.width > 2 and r.height > 2:
                    fills.append((round(r.x0, 2), round(r.y0, 2),
                                  round(r.x1, 2), round(r.y1, 2),
                                  tuple(round(c, 3) for c in p["fill"])))
            elif it[0] == "l":
                a, b = it[1], it[2]
                if abs(a.x - b.x) < 0.05 and 4.0 < abs(a.y - b.y) < 8.0:
                    vticks.append(round(a.x, 2))       # value-axis tick (below plot)
                elif abs(a.y - b.y) < 0.05 and 4.0 < abs(a.x - b.x) < 8.0:
                    hticks.append(round(a.y, 2))       # category tick (left of plot)

    vticks = sorted(set(vticks))
    hticks = sorted(set(hticks))
    if vticks:
        print(f"  plot_left={vticks[0]:.2f} (= sx+{vticks[0]-SX:.2f})  "
              f"plot_right={vticks[-1]:.2f}  div={len(vticks)-1}")
    if hticks:
        print(f"  plot_top={hticks[0]:.2f} (= sy+{hticks[0]-SY:.2f})  "
              f"plot_bot={hticks[-1]:.2f} (= sy+shh-{SY+SH-hticks[-1]:.2f})")

    for f in sorted(fills, key=lambda t: (t[1], t[0]))[:8]:
        print(f"  fill x[{f[0]:7.2f},{f[2]:7.2f}] y[{f[1]:7.2f},{f[3]:7.2f}] "
              f"w={f[2]-f[0]:6.2f} h={f[3]-f[1]:5.2f} {f[4]}")

    spans = []
    for blk in page.get_text("rawdict")["blocks"]:
        for ln in blk.get("lines", []):
            for sp in ln["spans"]:
                t = "".join(c["c"] for c in sp.get("chars", []))
                if t.strip():
                    spans.append((round(sp["origin"][1], 2), round(sp["origin"][0], 2),
                                  round(sp["bbox"][2], 2), round(sp["size"], 2),
                                  sp["font"], t.strip()))
    for s in sorted(spans):
        print(f"  span y={s[0]:7.2f} x0={s[1]:7.2f} x1={s[2]:7.2f} "
              f"sz={s[3]:5.2f} {s[4]:<26s} {s[5]!r}")
