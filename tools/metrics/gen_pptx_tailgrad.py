# -*- coding: utf-8 -*-
"""Probe: what gradient does PowerPoint paint on a callout tail outside the box?

252cbc8c s92's wedgeRoundRectCallouts take the theme's gradient fill
(fillRef idx=2) and their tails reach below / beside the shape box. Each arm
is a wedgeRectCallout with a black->white linear gradient at ang 0 or 90, its
tail pushed out of the box downward, rightward, or both. The PDF is
rasterised and sampled along the box's axis and down the tail, so the output
says whether the gradient spans the shape box (tail clamped to the end stop)
or the path's bounds (the box itself never reaches white).
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\tailgrad").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
# (angle deg, adj1, adj2)
ARMS = [(90, 0, 150000), (0, 0, 150000), (90, 150000, 0), (0, 150000, 0), (90, 120000, 140000), (45, 0, 150000)]
BOX = (120.0, 80.0)


def q(t: str, ns: str = A) -> str:
    return f"{{{ns}}}{t}"


def emu(pt: float) -> Emu:
    return Emu(int(round(pt * 12700)))


def origin(i: int) -> tuple[float, float]:
    return 40 + (i % 3) * 300, 40 + (i // 3) * 250


def build() -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = emu(960), emu(540)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    from pptx.enum.shapes import MSO_SHAPE
    for i, (ang, a1, a2) in enumerate(ARMS):
        x, y = origin(i)
        sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGULAR_CALLOUT, emu(x), emu(y), emu(BOX[0]), emu(BOX[1]))
        sp = sh._element
        sppr = sp.find(q("spPr", P))
        pg = sppr.find(q("prstGeom"))
        av = pg.find(q("avLst"))
        for c in list(av):
            av.remove(c)
        for n, v in (("adj1", a1), ("adj2", a2)):
            gd = etree.SubElement(av, q("gd"))
            gd.set("name", n)
            gd.set("fmla", f"val {v}")
        for tag in ("solidFill", "gradFill", "noFill"):
            for e in sppr.findall(q(tag)):
                sppr.remove(e)
        gf = etree.Element(q("gradFill"))
        gl = etree.SubElement(gf, q("gsLst"))
        for pos, col in ((0, "000000"), (100000, "FFFFFF")):
            gs = etree.SubElement(gl, q("gs"))
            gs.set("pos", str(pos))
            etree.SubElement(gs, q("srgbClr")).set("val", col)
        lin = etree.SubElement(gf, q("lin"))
        lin.set("ang", str(ang * 60000))
        lin.set("scaled", "0")
        pg.addnext(gf)
        ln = etree.Element(q("ln"))
        etree.SubElement(ln, q("noFill"))
        gf.addnext(ln)
        st = sp.find(q("style", P))
        if st is not None:
            sp.remove(st)
    dst = OUT / "tailgrad.pptx"
    prs.save(dst)
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    f = build()
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(f), WithWindow=False)
        pres.SaveAs(str(f.with_suffix(".pdf")), 32)
        pres.Close()
    finally:
        app.Quit()
    import pymupdf
    page = pymupdf.open(f.with_suffix(".pdf"))[0]
    dpi = 72
    pix = page.get_pixmap(dpi=dpi)
    def lum(x: float, y: float) -> int:
        px = pix.pixel(int(x), int(y))
        return round(sum(px[:3]) / 3)
    for i, (ang, a1, a2) in enumerate(ARMS):
        x, y = origin(i)
        w, h = BOX
        tx, ty = x + w / 2 + a1 / 100000 * w, y + h / 2 + a2 / 100000 * h
        box_v = [lum(x + w / 2, y + h * k) for k in (0.02, 0.5, 0.98)]
        box_h = [lum(x + w * k, y + h / 2) for k in (0.02, 0.5, 0.98)]
        # samples on the segment from the box centre-edge toward the tip
        tail = []
        for k in (0.3, 0.6, 0.85):
            sx = x + w / 2 + (tx - (x + w / 2)) * 1.0
            cx, cy = x + w / 2, y + h / 2
            px_, py_ = cx + (tx - cx) * (0.5 + 0.5 * k), cy + (ty - cy) * (0.5 + 0.5 * k)
            tail.append(((round(px_), round(py_)), lum(px_, py_)))
        print(f"ang={ang:3d} adj=({a1},{a2}) box_v(top,mid,bot)={box_v} box_h(l,mid,r)={box_h} tail={tail}")


if __name__ == "__main__":
    main()
