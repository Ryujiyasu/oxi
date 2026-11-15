# -*- coding: utf-8 -*-
"""Surface-chart geometry sweep (chart_surface_sweep): how PowerPoint fits the 3-D box.

Same construction as gen_pptx_chart_surface.py (a python-pptx LINE chart rewritten into
<c:surfaceChart>), one arm per slide, varying only the levers of the 3-D box fit:
frame width / height, category count, series count, value range. PowerPoint exports the
PDF; the axis ticks are vectors there (the mesh and the walls are a raster image), so
the projected box corners can be read back.

NEVER run this while the renderer is producing PNGs (`pptx_com_render_must_not_overlap`).

Run:  python tools/metrics/gen_pptx_chart_surface_sweep.py
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

sys.stdout.reconfigure(encoding="utf-8")

NAME = "chart_surface_sweep2" if "2" in sys.argv[1:] else "chart_surface_sweep"
DIR = os.path.abspath(os.path.join("pipeline_data", "pptx_probes", NAME))
OUT = os.path.join(DIR, NAME + ".pptx")
PDF = os.path.join(DIR, NAME + ".pdf")

BASE = [19.2, 21.4, 16.7, 22.0, 18.1, 20.3]
# (frame w, frame h, n categories, n series, value scale)
ARMS = [
    (396, 288, 4, 2, 1.0),
    (300, 288, 4, 2, 1.0),
    (500, 288, 4, 2, 1.0),
    (600, 288, 4, 2, 1.0),
    (396, 200, 4, 2, 1.0),
    (396, 400, 4, 2, 1.0),
    (396, 288, 3, 2, 1.0),
    (396, 288, 6, 2, 1.0),
    (396, 288, 4, 3, 1.0),
    (396, 288, 4, 4, 1.0),
    (396, 288, 6, 4, 1.0),
    (396, 288, 4, 2, 4.0),
    (600, 400, 4, 3, 1.0),
    (300, 200, 3, 3, 1.0),
]


# sweep 2 (`... 2`): depth (series/categories) finely at one plot, then the aspect at D=1
if "2" in sys.argv[1:]:
    ARMS = [(396, 288, nc, 2, 1.0) for nc in (2, 3, 5, 8, 10)] +            [(396, 288, 4, ns, 1.0) for ns in (5, 6)] +            [(396, hh, 4, 4, 1.0) for hh in (200, 350, 400)] +            [(ww, 288, 4, 4, 1.0) for ww in (300, 500)] +            [(396, 288, 2, 2, 1.0), (396, 288, 3, 3, 1.0), (396, 288, 6, 6, 1.0)]
    BASE = BASE + [17.5, 19.9, 21.1, 18.8]


def pt(v):
    return Emu(int(round(v * 12700)))


def build():
    prs = Presentation()
    prs.slide_width = pt(720)
    prs.slide_height = pt(540)
    for w, h, nc, ns, k in ARMS:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        cd = CategoryChartData()
        cd.categories = ["Q%d" % (i + 1) for i in range(nc)]
        for s in range(ns):
            cd.add_series("Ser%d" % (s + 1),
                          [round(BASE[i] * k * (1 - 0.18 * s), 1) for i in range(nc)])
        gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE, pt(72), pt(72), pt(w), pt(h), cd)
        gf.chart.has_title = False
        gf.chart.has_legend = False
    os.makedirs(DIR, exist_ok=True)
    prs.save(OUT)


def ser_ax(val_ax_id):
    return ('<c:serAx><c:axId val="900000003"/><c:scaling>'
            '<c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
            '<c:axPos val="b"/><c:majorTickMark val="out"/>'
            '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
            '<c:crossAx val="%s"/></c:serAx>' % val_ax_id)


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
            x = x.replace("</c:valAx>", "</c:valAx>" + ser_ax(vid), 1)
            data = x.encode("utf-8")
        zout.writestr(it, data)
    zin.close()
    zout.close()
    shutil.move(tmp, OUT)


def export():
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(OUT, WithWindow=False)
    pres.SaveAs(PDF, 32)
    pres.Close()


if __name__ == "__main__":
    build()
    surgery()
    export()
    with open(os.path.join(DIR, "arms.json"), "w") as f:
        json.dump(ARMS, f)
    print("wrote", OUT, PDF)
