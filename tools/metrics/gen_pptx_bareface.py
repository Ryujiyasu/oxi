# -*- coding: utf-8 -*-
"""Probe: which face and size does PowerPoint give a run whose rPr states
neither (`<a:rPr lang="en-US"/>`) when its siblings state both?

d44 slide 5 (Google Slides export) has such a run holding one space between
Arimo 22pt runs in a text box; the presentation defaultTextStyle says Arial
14pt. Each arm is a text box: [Times New Roman 28pt "Ab"] + [bare run] +
[Times New Roman 28pt "Ab"]; the PDF's chars say which face and size the bare
run's glyphs got and how wide its space is.
  t1  bare run "xyz"         (visible glyphs)
  t2  bare run " "           (space only)
  t3  bare run "xyz", siblings with no face but sz 2800
  p1  same as t1 in a body placeholder (master body style)
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\bareface").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def q(t: str) -> str:
    return f"{{{A}}}{t}"


def fill(body, bare: str, sib_face: bool) -> None:
    for pel in body.findall(q("p")):
        body.remove(pel)
    p = etree.SubElement(body, q("p"))
    for k, txt in enumerate(("Ab", bare, "Ab")):
        r = etree.SubElement(p, q("r"))
        rpr = etree.SubElement(r, q("rPr"))
        rpr.set("lang", "en-US")
        if k != 1:
            rpr.set("sz", "2800")
            if sib_face:
                etree.SubElement(rpr, q("latin")).set("typeface", "Times New Roman")
        etree.SubElement(r, q("t")).text = txt


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for k, (bare, face) in enumerate((("xyz", True), (" ", True), ("xyz", False))):
        box = slide.shapes.add_textbox(Emu(457200), Emu(457200 + k * 1000000), Emu(6400000), Emu(800000))
        fill(box.text_frame._txBody, bare, face)
    s2 = prs.slides.add_slide(prs.slide_layouts[1])
    fill(s2.placeholders[1].text_frame._txBody, "xyz", True)
    f = OUT / "bareface.pptx"
    prs.save(f)
    import pymupdf
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(f), WithWindow=False)
        pdf = f.with_suffix(".pdf")
        pres.SaveAs(str(pdf), 32)
        pres.Close()
    finally:
        app.Quit()
    for pi, page in enumerate(pymupdf.open(pdf)):
        for b in page.get_text("rawdict")["blocks"]:
            for line in b.get("lines", []):
                for s in line["spans"]:
                    xs = [round(c["bbox"][0], 2) for c in s["chars"]]
                    print(pi + 1, s["font"], round(s["size"], 2), repr("".join(c["c"] for c in s["chars"])), xs)


if __name__ == "__main__":
    main()
