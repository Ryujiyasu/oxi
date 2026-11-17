# -*- coding: utf-8 -*-
"""Probe: when does an a:br's own rPr/@sz set a line's height?

252cbc s104: "14pt text <br sz=1600/> 16pt url" keeps the first line at 14pt,
while abde7 s1: "36pt text <br sz=1400/><br sz=1400/> 36pt text" draws the
empty middle line at 14pt. Each arm is one zero-inset box; the PDF baselines
of its text lines are printed so the line pitch can be read off.
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\brsz").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# each arm: list of ("t", size, text) / ("br", size or None)
ARMS = {
    "text14_br40_text14": [("t", 14, "Abc"), ("br", 40), ("t", 14, "Def")],
    "text14_br8_text14": [("t", 14, "Abc"), ("br", 8), ("t", 14, "Def")],
    "text14_brNone_text14": [("t", 14, "Abc"), ("br", None), ("t", 14, "Def")],
    "text14_br14_br40_text14": [("t", 14, "Abc"), ("br", 14), ("br", 40), ("t", 14, "Def")],
    "text14_br40_br14_text14": [("t", 14, "Abc"), ("br", 40), ("br", 14), ("t", 14, "Def")],
    "text40_br14_br14_text40": [("t", 40, "Abc"), ("br", 14), ("br", 14), ("t", 40, "Def")],
    "text14_brNone_brNone_text14": [("t", 14, "Abc"), ("br", None), ("br", None), ("t", 14, "Def")],
    "br40_text14": [("br", 40), ("t", 14, "Def")],
    "text14_text14": [("t", 14, "Abc"), ("p",), ("t", 14, "Def")],
}


def q(t: str) -> str:
    return f"{{{A}}}{t}"


def emu(pt: float) -> Emu:
    return Emu(int(round(pt * 12700)))


def build() -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = emu(960), emu(540)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for k, (name, items) in enumerate(ARMS.items()):
        box = slide.shapes.add_textbox(emu(10 + k * 100), emu(10), emu(95), emu(500))
        body = box.text_frame._txBody
        bp = body.find(q("bodyPr"))
        for a in ("lIns", "rIns", "tIns", "bIns"):
            bp.set(a, "0")
        for pel in body.findall(q("p")):
            body.remove(pel)
        p = etree.SubElement(body, q("p"))
        for it in items:
            if it[0] == "p":
                p = etree.SubElement(body, q("p"))
                continue
            el = etree.SubElement(p, q("r" if it[0] == "t" else "br"))
            rpr = etree.SubElement(el, q("rPr"))
            rpr.set("lang", "en-US")
            if it[1] is not None:
                rpr.set("sz", str(it[1] * 100))
            etree.SubElement(rpr, q("latin")).set("typeface", "Arial")
            if it[0] == "t":
                etree.SubElement(el, q("t")).text = it[2]
    dst = OUT / "brsz.pptx"
    prs.save(dst)
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    f = build()
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
    page = pymupdf.open(pdf)[0]
    for k, name in enumerate(ARMS):
        x0 = 10 + k * 100
        ys = []
        for b in page.get_text("dict")["blocks"]:
            for line in b.get("lines", []):
                for sp in line["spans"]:
                    ox, oy = sp["origin"]
                    if x0 - 1 <= ox <= x0 + 95 and sp["text"].strip():
                        ys.append((round(oy, 2), sp["text"].strip(), sp["size"]))
        print(name.ljust(30), sorted(set(ys)))


if __name__ == "__main__":
    main()
