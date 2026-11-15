# -*- coding: utf-8 -*-
"""Surface-chart axis auto-settings straight from PowerPoint's object model.

Builds surface charts over a grid of (categories, series, frame size), opens the deck in
PowerPoint and reads the series axis `TickLabelSpacing` (the label thinning interval) and
the value axis `MajorUnit` / `MinimumScale` / `MaximumScale`, plus the plot-area inside
height. These agree with what PowerPoint draws in its PDF (chart_surface_sweep5: all 16
arms), so the rule behind them can be fitted from many more arms than a PDF sweep gives.

NEVER run this while the renderer is producing PNGs (`pptx_com_render_must_not_overlap`).

Run:  python tools/metrics/com_surface_axis_sweep.py   (writes .../com_surface_axis/rows.json)
"""
import json
import os
import re
import shutil
import sys
import zipfile

import win32com.client
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Emu

DIR = os.path.abspath(os.path.join("pipeline_data", "pptx_probes", "com_surface_axis"))
OUT = os.path.join(DIR, "com_surface_axis.pptx")
BASE = [19.2, 21.4, 16.7, 22.0, 18.1, 20.3, 17.5, 19.9, 21.1, 18.8]
ARMS = [(w, h, nc, ns) for (w, h) in ((396, 288), (300, 240), (560, 360))
        for nc in range(2, 11) for ns in range(2, 11)]
# `place`: the vertical-placement grid (sizes x depth), read for the axis boxes
if "place" in sys.argv[1:]:
    ARMS = [(w, h, nc, ns) for (w, h) in ((300, 200), (300, 300), (400, 250), (400, 350),
                                         (500, 300), (600, 250), (600, 400))
            for nc in range(2, 9) for ns in range(2, 9)]
    OUT = OUT.replace("com_surface_axis.pptx", "com_surface_place.pptx")


def pt(v):
    return Emu(int(round(v * 12700)))


def build():
    prs = Presentation()
    prs.slide_width = pt(720)
    prs.slide_height = pt(540)
    for w, h, nc, ns in ARMS:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        cd = CategoryChartData()
        cd.categories = ["Q%d" % (i + 1) for i in range(nc)]
        for s in range(ns):
            cd.add_series("Ser%d" % (s + 1), [round(BASE[i] * (1 - 0.08 * s), 1) for i in range(nc)])
        gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE, pt(72), pt(72), pt(w), pt(h), cd)
        gf.chart.has_title = False
        gf.chart.has_legend = False
    os.makedirs(DIR, exist_ok=True)
    prs.save(OUT)


def surgery():
    tmp = OUT + ".tmp"
    zin = zipfile.ZipFile(OUT)
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        data = zin.read(it.filename)
        if re.match(r"ppt/charts/chart(\d+)\.xml$", it.filename):
            x = data.decode("utf-8")
            x = re.sub(r"<c:grouping[^>]*/>", "", x)
            x = re.sub(r"<c:varyColors[^>]*/>", "", x)
            x = re.sub(r"<c:marker>.*?</c:marker>", "", x, flags=re.S)
            x = re.sub(r"<c:marker[^>]*/>", "", x)
            x = re.sub(r"<c:smooth[^>]*/>", "", x)
            x = x.replace("<c:lineChart>", '<c:surfaceChart><c:wireframe val="0"/>')
            x = x.replace("</c:lineChart>", "</c:surfaceChart>")
            x = re.sub(r'(<c:axId val="\d+"/><c:axId val="\d+"/>)(</c:surfaceChart>)',
                       r'\1<c:axId val="900000003"/>\2', x)
            vid = re.search(r'<c:valAx><c:axId val="(\d+)"/>', x).group(1)
            x = x.replace("</c:valAx>", '</c:valAx><c:serAx><c:axId val="900000003"/><c:scaling>'
                          '<c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
                          '<c:axPos val="b"/><c:majorTickMark val="out"/>'
                          '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
                          '<c:crossAx val="%s"/></c:serAx>' % vid, 1)
            data = x.encode("utf-8")
        zout.writestr(it, data)
    zin.close()
    zout.close()
    shutil.move(tmp, OUT)


def read():
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(OUT, WithWindow=False)
    rows = []
    for (w, h, nc, ns), slide in zip(ARMS, pres.Slides):
        for sh in slide.Shapes:
            if not sh.HasChart:
                continue
            c = sh.Chart
            a3, a2 = c.Axes(3), c.Axes(2)
            pa = c.PlotArea
            boxes = {}
            for ai in (1, 2, 3):
                ax = c.Axes(ai)
                boxes["ax%d" % ai] = (ax.Left, ax.Top, ax.Width, ax.Height)
            rows.append(dict(w=w, h=h, nc=nc, ns=ns, k=a3.TickLabelSpacing, major=a2.MajorUnit, **boxes,
                             vmin=a2.MinimumScale, vmax=a2.MaximumScale,
                             in_top=pa.InsideTop, in_h=pa.InsideHeight,
                             in_left=pa.InsideLeft, in_w=pa.InsideWidth))
    pres.Close()
    return rows


if __name__ == "__main__":
    build()
    surgery()
    rows = read()
    name = "rows_place.json" if "place" in sys.argv[1:] else "rows.json"
    json.dump(rows, open(os.path.join(DIR, name), "w"), indent=0)
    print("arms", len(rows))
