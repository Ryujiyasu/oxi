# -*- coding: utf-8 -*-
"""Probe: how tall is an empty paragraph whose a:endParaRPr states no sz?

`emptypara` (2026-08-18) settled endParaRPr/@sz (wins) and a missing
endParaRPr (the PRECEDING paragraph's size). abde73e7 slide 7 has the third
case -- `<a:endParaRPr lang=.. kern="0"/>` with no sz between a 28pt line and
an 18pt level -- and PowerPoint leaves 14.4pt less than Oxi's
previous-size rule. Each arm is a text box: 28pt line / empty paragraph /
"X" at the level size (defaultTextStyle lvl1 sz=1800); COM reads the third
paragraph's BoundTop.
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\emptymark").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
ARMS = [
    ("mark_nosz", "nosz"),
    ("no_mark", None),
    ("mark_10", "1000"),
    ("mark_28", "2800"),
]


def q(t: str) -> str:
    return f"{{{A}}}{t}"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for k, (label, mark) in enumerate(ARMS):
        box = slide.shapes.add_textbox(Emu(457200 + k * 2000000), Emu(457200), Emu(1800000), Emu(3000000))
        body = box.text_frame._txBody
        for pel in body.findall(q("p")):
            body.remove(pel)
        p1 = etree.SubElement(body, q("p"))
        r = etree.SubElement(p1, q("r"))
        etree.SubElement(r, q("rPr")).set("sz", "2800")
        etree.SubElement(r, q("t")).text = "Big"
        p2 = etree.SubElement(body, q("p"))
        if mark is not None:
            e = etree.SubElement(p2, q("endParaRPr"))
            e.set("lang", "en-US")
            if mark != "nosz":
                e.set("sz", mark)
        p3 = etree.SubElement(body, q("p"))
        r = etree.SubElement(p3, q("r"))
        etree.SubElement(r, q("rPr")).set("lang", "en-US")
        etree.SubElement(r, q("t")).text = "X"
    dst = OUT / "emptymark.pptx"
    prs.save(dst)
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(dst), WithWindow=False)
        for k, (label, _) in enumerate(ARMS, start=1):
            tr = pres.Slides(1).Shapes(k).TextFrame.TextRange
            t1 = tr.Paragraphs(1).BoundTop
            print(f"{label:10s} p2_top-p1_top={tr.Paragraphs(2).BoundTop - t1:7.2f}"
                  f" p3_top-p1_top={tr.Paragraphs(3).BoundTop - t1:7.2f}"
                  f" p2_h={tr.Paragraphs(2).BoundHeight:6.2f}")
        pres.SaveAs(str(dst.with_suffix(".pdf")), 32)
        pres.Close()
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
