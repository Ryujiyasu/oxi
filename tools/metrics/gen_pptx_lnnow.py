# -*- coding: utf-8 -*-
"""Probe: how thick does PowerPoint draw an `a:ln` that states a colour but no @w?

abde73e7 s1's QR picture carries `<a:ln><a:solidFill>..</a:solidFill></a:ln>`
with no width and no p:style, and PowerPoint frames it; Oxi drew nothing
(border width None). Arms: a rect autoshape and a picture, each with a red
`a:ln` and no @w, and controls at w=9525 / 12700. Output: the stroke width
PowerPoint writes into its PDF for each red path (pymupdf drawings).
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\lnnow").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
ARMS = [("rect", None), ("pic", None), ("rect", 9525), ("pic", 9525), ("rect", 12700)]


def q(t: str, ns: str = A) -> str:
    return f"{{{ns}}}{t}"


def emu(pt: float) -> Emu:
    return Emu(int(round(pt * 12700)))


def png() -> io.BytesIO:
    from PIL import Image
    b = io.BytesIO()
    Image.new("RGB", (40, 40), (230, 230, 230)).save(b, "PNG")
    b.seek(0)
    return b


def build() -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = emu(960), emu(540)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    from pptx.enum.shapes import MSO_SHAPE
    for i, (kind, w) in enumerate(ARMS):
        x, y = 40 + i * 180, 100
        if kind == "rect":
            sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, emu(x), emu(y), emu(120), emu(120))
            st = sh._element.find(q("style", P))
            if st is not None:
                sh._element.remove(st)
            sppr = sh._element.find(q("spPr", P))
            etree.SubElement(sppr, q("noFill"))
        else:
            sh = slide.shapes.add_picture(png(), emu(x), emu(y), emu(120), emu(120))
            sppr = sh._element.find(q("spPr", P))
        ln = etree.SubElement(sppr, q("ln"))
        if w is not None:
            ln.set("w", str(w))
        sf = etree.SubElement(ln, q("solidFill"))
        etree.SubElement(sf, q("srgbClr")).set("val", "FF0000")
    dst = OUT / "lnnow.pptx"
    prs.save(dst)
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    f = build()
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(f), WithWindow=False)
        weights = []
        for sh in pres.Slides(1).Shapes:
            weights.append((sh.Line.Visible, round(sh.Line.Weight, 3)))
        pres.SaveAs(str(f.with_suffix(".pdf")), 32)
        pres.Close()
    finally:
        app.Quit()
    import pymupdf
    page = pymupdf.open(f.with_suffix(".pdf"))[0]
    found = {}
    for d in page.get_drawings():
        c = d.get("color")
        if c and c[0] > 0.9 and c[1] < 0.1:
            i = int((d["rect"].x0 - 40 + 30) // 180)
            found.setdefault(i, set()).add(round(d.get("width") or 0, 3))
    for i, (kind, w) in enumerate(ARMS):
        print(kind, w, "com", weights[i], "pdf widths", sorted(found.get(i, [])))


if __name__ == "__main__":
    main()
