# -*- coding: utf-8 -*-
"""Probe: how far does a PARAGRAPH's own `a:spcBef/a:spcPct` (and spcAft) move
the next line?

abde73e7 puts spcBef 20% on every paragraph of its text boxes and Oxi added
nothing (only spcPts was read at paragraph level). The master-level rule is
`pct * size * 1.2 * lnSpc`. Arms: size 20 / 40, lnSpc 100 / 150, spcBef
0 / 20 / 50 / 100 %, plus spcAft 50 %.
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\paraspcpct").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
# (size pt, lnSpc %, [(spcBef %, spcAft %) per paragraph])
ARMS = [
    (20, 100, [(0, 0), (20, 0), (50, 0), (100, 0), (0, 50), (0, 0)]),
    (40, 100, [(0, 0), (20, 0), (50, 0), (0, 50), (0, 0)]),
    (20, 150, [(0, 0), (20, 0), (50, 0), (100, 0), (0, 50), (0, 0)]),
]


def q(tag: str) -> str:
    return f"{{{A}}}{tag}"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    blank = prs.slide_layouts[6]
    for size, lnspc, paras in ARMS:
        slide = prs.slides.add_slide(blank)
        box = slide.shapes.add_textbox(Emu(914400), Emu(457200), Emu(6400800), Emu(5000000))
        body = box.text_frame._txBody
        for pel in body.findall(q("p")):
            body.remove(pel)
        for k, (bef, aft) in enumerate(paras):
            p = etree.SubElement(body, q("p"))
            ppr = etree.SubElement(p, q("pPr"))
            ln = etree.SubElement(ppr, q("lnSpc"))
            etree.SubElement(ln, q("spcPct")).set("val", str(lnspc * 1000))
            if bef:
                sb = etree.SubElement(ppr, q("spcBef"))
                etree.SubElement(sb, q("spcPct")).set("val", str(bef * 1000))
            if aft:
                sa = etree.SubElement(ppr, q("spcAft"))
                etree.SubElement(sa, q("spcPct")).set("val", str(aft * 1000))
            r = etree.SubElement(p, q("r"))
            rpr = etree.SubElement(r, q("rPr"))
            rpr.set("lang", "en-US")
            rpr.set("sz", str(size * 100))
            etree.SubElement(rpr, q("latin")).set("typeface", "Arial")
            etree.SubElement(r, q("t")).text = f"Line {k} bef {bef} aft {aft}"
    prs.save(OUT / "paraspcpct.pptx")
    print("wrote", OUT / "paraspcpct.pptx")


if __name__ == "__main__":
    main()
