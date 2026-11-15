# -*- coding: utf-8 -*-
"""Probe: how PowerPoint draws every `a:prstDash` preset, on lines, closed outlines,
group members, connectors, round caps and table cell borders.

PowerPoint authors the shapes (each one dashed `dash` through COM); the slide XML is then
patched so each shape carries its own preset, and PowerPoint itself re-opens the patched
file and exports the PDF that is the truth.

NEVER run this while the renderer is producing PNGs (`pptx_com_render_must_not_overlap`).

Usage: python tools/metrics/gen_pptx_dash.py
"""
import re
import shutil
import zipfile
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_probes" / "dash"
OUT.mkdir(parents=True, exist_ok=True)
RAW = OUT / "raw.pptx"
PPTX = OUT / "dash.pptx"
PDF = OUT / "dash.pdf"

PRESETS = ["dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot",
           "sysDash", "sysDot", "sysDashDot", "sysDashDotDot", "solid"]

app = win32com.client.Dispatch("PowerPoint.Application")
pres = app.Presentations.Add(WithWindow=False)
pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight = 720, 540
BLANK = 12  # ppLayoutBlank


def dashed(shape, width, color=0x000000):
    shape.Line.Visible = True
    shape.Line.ForeColor.RGB = color
    shape.Line.Weight = width
    shape.Line.DashStyle = 4  # msoLineDash -> <a:prstDash val="dash"/>, patched below


# slides 1-3: horizontal lines, one preset per row, at 0.75 / 2 / 4.5pt
for w in (0.75, 2.0, 4.5):
    s = pres.Slides.Add(pres.Slides.Count + 1, BLANK)
    for i, _ in enumerate(PRESETS):
        y = 40 + i * 44
        ln = s.Shapes.AddLine(60, y, 660, y)
        dashed(ln, w)

# slide 4: closed outlines (rect / ellipse / roundRect) per preset, 2pt
s = pres.Slides.Add(pres.Slides.Count + 1, BLANK)
for i, _ in enumerate(PRESETS):
    col, row = i % 4, i // 4
    sh = s.Shapes.AddShape(1 if row != 1 else 9, 40 + col * 170, 40 + row * 160, 140, 120)
    sh.Fill.Visible = False
    dashed(sh, 2.0)

# slide 5: group members, connectors, round-cap dots
s = pres.Slides.Add(pres.Slides.Count + 1, BLANK)
a = s.Shapes.AddShape(1, 60, 60, 160, 100)
b = s.Shapes.AddShape(9, 260, 60, 160, 100)
c = s.Shapes.AddLine(60, 200, 420, 200)
for sh in (a, b, c):
    sh.Fill.Visible = False
    dashed(sh, 2.0)
s.Shapes.Range([a.Name, b.Name, c.Name]).Group()
e = s.Shapes.AddConnector(2, 460, 60, 660, 200)  # msoConnectorElbow
dashed(e, 2.0)
f = s.Shapes.AddConnector(1, 460, 260, 660, 400)  # straight
dashed(f, 3.0)
g = s.Shapes.AddLine(60, 300, 420, 300)
dashed(g, 6.0)
h = s.Shapes.AddLine(60, 360, 420, 360)
dashed(h, 6.0)
k = s.Shapes.AddLine(60, 420, 420, 420)
dashed(k, 6.0)

# slide 6: a table whose cell borders are dashed
s = pres.Slides.Add(pres.Slides.Count + 1, BLANK)
t = s.Shapes.AddTable(3, 3, 60, 60, 600, 240).Table
for r in range(1, 4):
    for cc in range(1, 4):
        for side in (1, 2, 3, 4):  # top left bottom right
            bd = t.Cell(r, cc).Borders(side)
            bd.Visible = True
            bd.Weight = 2.0
            bd.ForeColor.RGB = 0
            bd.DashStyle = 4

pres.SaveAs(str(RAW))
pres.Close()

# patch: one preset per dashed shape, in document order
SLIDE_PLAN = {
    1: PRESETS, 2: PRESETS, 3: PRESETS, 4: PRESETS,
    # group rect / group roundRect / group line / elbow / straight / cap rnd dot / cap rnd sysDot / flat lgDashDot
    5: ["dot", "lgDashDot", "sysDashDot", "dashDot", "lgDash", "dot", "sysDot", "lgDashDot"],
}
shutil.copy(RAW, PPTX)
zin = zipfile.ZipFile(RAW)
tmp = OUT / "tmp.pptx"
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        m = re.match(r"ppt/slides/slide(\d+)\.xml$", item.filename)
        if m:
            n = int(m.group(1))
            x = data.decode("utf-8")
            if n in SLIDE_PLAN:
                plan = iter(SLIDE_PLAN[n])
                x = re.sub(r'<a:prstDash val="dash"/>',
                           lambda _: f'<a:prstDash val="{next(plan)}"/>', x)
                if n == 5:
                    # the two 6pt round-cap arms: cap="rnd" on the 6th and 7th <a:ln
                    idx = [mm.start() for mm in re.finditer(r"<a:ln w=", x)]
                    for j in (6, 5):
                        if j < len(idx):
                            p = idx[j]
                            x = x[:p] + '<a:ln cap="rnd" w=' + x[p + len("<a:ln w="):]
            if n == 6:
                cyc = iter(["dot", "dash", "lgDashDot", "sysDash"] * 40)
                x = re.sub(r'<a:prstDash val="dash"/>',
                           lambda _: f'<a:prstDash val="{next(cyc)}"/>', x)
            data = x.encode("utf-8")
        zout.writestr(item, data)
zin.close()
shutil.move(tmp, PPTX)

pres = app.Presentations.Open(str(PPTX), WithWindow=False)
pres.SaveAs(str(PDF), 32)
pres.Close()
print("wrote", PPTX, PDF)
