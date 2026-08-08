# -*- coding: utf-8 -*-
"""Export chart_bar_axis via PowerPoint COM and read the horizontal value-axis
tick count + labels per slide."""
import sys, os

sys.stdout.reconfigure(encoding="utf-8")
import win32com.client
import fitz

BASE = r"pipeline_data\pptx_probes\chart_bar_axis"
PPTX = os.path.abspath(os.path.join(BASE, "chart_bar_axis.pptx"))
PDF = os.path.abspath(os.path.join(BASE, "chart_bar_axis.pdf"))

if not os.path.exists(PDF) or os.path.getmtime(PDF) < os.path.getmtime(PPTX):
    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = ppt.Presentations.Open(PPTX, WithWindow=False)
        pres.SaveAs(PDF, 32)
        pres.Close()
    finally:
        ppt.Quit()
    print("exported:", PDF)

doc = fitz.open(PDF)
print(f"{'pg':>3} {'plot_l':>8} {'plot_r':>8} {'plot_w':>8} {'div':>4} {'labels'}")
for pno in range(doc.page_count):
    page = doc[pno]
    xs = []
    for p in page.get_drawings():
        for it in p["items"]:
            if it[0] == "l":
                a, b = it[1], it[2]
                # short vertical tick below the value axis
                if abs(a.x - b.x) < 0.05 and 4.0 < abs(a.y - b.y) < 8.0:
                    xs.append(round(a.x, 2))
    xs = sorted(set(xs))
    labels = []
    d = page.get_text("rawdict")
    for blk in d["blocks"]:
        for ln in blk.get("lines", []):
            for sp in ln["spans"]:
                t = "".join(c["c"] for c in sp.get("chars", []))
                if t.strip():
                    labels.append((round(sp["origin"][1], 2), t.strip()))
    # value labels sit on the lowest baseline (the axis band)
    if labels:
        ymax = max(y for y, _ in labels)
        vlabels = [t for y, t in labels if abs(y - ymax) < 1.0]
    else:
        vlabels = []
    div = len(xs) - 1 if xs else 0
    pl = xs[0] if xs else 0.0
    pr = xs[-1] if xs else 0.0
    print(f"{pno+1:>3} {pl:8.2f} {pr:8.2f} {pr-pl:8.2f} {div:>4} {vlabels}")
