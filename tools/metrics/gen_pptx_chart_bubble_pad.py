# -*- coding: utf-8 -*-
"""Bubble-chart axis padding sweep (chart_bubble_pad): when does PowerPoint widen the X
and Y axes so the bubbles fit, and by how much?

One bubble chart per slide (python-pptx, then `c:bubbleScale` patched in), varying the
bubble scale and where the data sits on X. PowerPoint exports the PDF; the axis labels
are read back as text.

NEVER run this while the renderer is producing PNGs (`pptx_com_render_must_not_overlap`).

Run:  python tools/metrics/gen_pptx_chart_bubble_pad.py
"""
import json
import os
import re
import shutil
import sys
import zipfile

import win32com.client
from pptx import Presentation
from pptx.chart.data import BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Emu

sys.stdout.reconfigure(encoding="utf-8")
DIR = os.path.abspath(os.path.join("pipeline_data", "pptx_probes", "chart_bubble_pad"))
OUT = os.path.join(DIR, "chart_bubble_pad.pptx")
PDF = os.path.join(DIR, "chart_bubble_pad.pdf")

# (bubble scale, x offset, frame w, frame h)
ARMS = [(s, 0, 396, 288) for s in (150, 200, 225, 250, 275, 300)] + \
       [(300, dx, 396, 288) for dx in (1, 2, 5)] + \
       [(300, 0, 600, 288), (300, 0, 396, 400), (200, 1, 600, 400)]


def pt(v):
    return Emu(int(round(v * 12700)))


def build():
    prs = Presentation()
    prs.slide_width = pt(720)
    prs.slide_height = pt(540)
    for scale, dx, w, h in ARMS:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        cd = BubbleChartData()
        s = cd.add_series("Ser1")
        for x, y, z in ((1, 19.2, 1), (2, 21.4, 2), (3, 16.7, 4)):
            s.add_data_point(x + dx, y, z)
        gf = slide.shapes.add_chart(XL_CHART_TYPE.BUBBLE, pt(72), pt(72), pt(w), pt(h), cd)
        gf.chart.has_legend = False
    os.makedirs(DIR, exist_ok=True)
    prs.save(OUT)


def patch():
    tmp = OUT + ".tmp"
    zin = zipfile.ZipFile(OUT)
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        data = zin.read(it.filename)
        m = re.match(r"ppt/charts/chart(\d+)\.xml$", it.filename)
        if m:
            scale = ARMS[int(m.group(1)) - 1][0]
            x = data.decode("utf-8")
            x = re.sub(r'<c:bubbleScale val="\d+"/>', "", x)
            # CT_BubbleChart order: ... dLbls?, bubble3D?, bubbleScale?, showNegBubbles?, ...
            anchor = "<c:showNegBubbles" if "<c:showNegBubbles" in x else "<c:axId "
            x = x.replace(anchor, '<c:bubbleScale val="%d"/>' % scale + anchor, 1)
            data = x.encode("utf-8")
        zout.writestr(it, data)
    zin.close()
    zout.close()
    shutil.move(tmp, OUT)


if __name__ == "__main__":
    build()
    patch()
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(OUT, WithWindow=False)
    pres.SaveAs(PDF, 32)
    pres.Close()
    json.dump(ARMS, open(os.path.join(DIR, "arms.json"), "w"))
    print("wrote", OUT, PDF)
