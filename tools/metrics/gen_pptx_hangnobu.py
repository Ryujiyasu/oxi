# -*- coding: utf-8 -*-
"""Probe: where does line 0 of a NO-bullet paragraph with a negative indent start?

Oxi places it at max(marL, -indent) -- the hanging-bullet rule -- so with
marL > -indent it lands on marL, level with the continuations. 252cbc8c slide 41
(marL 141.7pt, indent -124.7pt, buNone) has PowerPoint draw line 0 at
marL + indent = 17pt. The arms separate the two readings: marL > -indent,
marL == -indent, marL < -indent, each with and without a bullet.
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\hangnobu").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
TEXT = ("alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu "
        "nu xi omicron pi rho sigma tau upsilon phi chi psi omega")
# (label, marL pt, indent pt, bullet)
ARMS = [
    ("m72_i-54_nobu", 72, -54, None),
    ("m72_i-54_bu", 72, -54, "•"),
    ("m36_i-36_nobu", 36, -36, None),
    ("m18_i-36_nobu", 18, -36, None),
    ("m18_i-36_bu", 18, -36, "•"),
    ("m108_i-90_nobu", 108, -90, None),
    ("m72_i-18_nobu", 72, -18, None),
]


def q(tag: str) -> str:
    return f"{{{A}}}{tag}"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    blank = prs.slide_layouts[6]
    for label, mar_l, indent, bullet in ARMS:
        slide = prs.slides.add_slide(blank)
        box = slide.shapes.add_textbox(Emu(914400), Emu(1200000), Emu(4572000), Emu(3000000))
        tf = box.text_frame
        tf.word_wrap = True
        body = tf._txBody
        bp = body.find(q("bodyPr"))
        for k in ("lIns", "rIns"):
            bp.set(k, "0")
        for pel in body.findall(q("p")):
            body.remove(pel)
        p = etree.SubElement(body, q("p"))
        ppr = etree.SubElement(p, q("pPr"))
        ppr.set("marL", str(int(mar_l * 12700)))
        ppr.set("indent", str(int(indent * 12700)))
        if bullet:
            etree.SubElement(ppr, q("buFont")).set("typeface", "Arial")
            etree.SubElement(ppr, q("buChar")).set("char", bullet)
        else:
            etree.SubElement(ppr, q("buNone"))
        r = etree.SubElement(p, q("r"))
        rpr = etree.SubElement(r, q("rPr"))
        rpr.set("lang", "en-US")
        rpr.set("sz", "1400")
        etree.SubElement(rpr, q("latin")).set("typeface", "Arial")
        etree.SubElement(r, q("t")).text = TEXT
    prs.save(OUT / "hangnobu.pptx")
    print(f"wrote {OUT / 'hangnobu.pptx'}  ({len(ARMS)} arms)")


if __name__ == "__main__":
    main()
