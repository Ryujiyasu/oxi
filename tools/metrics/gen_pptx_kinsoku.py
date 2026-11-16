# -*- coding: utf-8 -*-
"""Probe: which characters may START a line in PowerPoint (Japanese), with the
paragraph's default kinsoku and with `eaLnBrk="0"`?

252cbc8c s45 breaks 「一つのデ／ータポータル」: the long vowel mark starts a line,
which Oxi's line-start-prohibited list forbids. Each arm is a text box exactly
five ideographs wide (20pt, zero insets) holding 「あああああ」+X+「あいう」;
if X may start a line the first line is the five あ, else PowerPoint pulls
something back (or pushes X in). The PDF lines say which.
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\kinsoku").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHARS = "ーッャァっゃぁ々・、。）」ゝヽ〜！？％"


def q(t: str) -> str:
    return f"{{{A}}}{t}"


def build(ea_ln_brk: str | None) -> Path:
    prs = Presentation()
    per = 6
    for s0 in range(0, len(CHARS), per):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for k, ch in enumerate(CHARS[s0:s0 + per]):
            box = slide.shapes.add_textbox(Emu(457200 + (k % 3) * 2700000), Emu(457200 + (k // 3) * 2600000),
                                           Emu(Pt(100.4).emu), Emu(Pt(150).emu))
            body = box.text_frame._txBody
            bp = body.find(q("bodyPr"))
            for a in ("lIns", "rIns", "tIns", "bIns"):
                bp.set(a, "0")
            bp.set("wrap", "square")
            for pel in body.findall(q("p")):
                body.remove(pel)
            p = etree.SubElement(body, q("p"))
            if ea_ln_brk is not None:
                etree.SubElement(p, q("pPr")).set("eaLnBrk", ea_ln_brk)
            r = etree.SubElement(p, q("r"))
            rpr = etree.SubElement(r, q("rPr"))
            rpr.set("lang", "ja-JP")
            rpr.set("sz", "2000")
            etree.SubElement(rpr, q("latin")).set("typeface", "ＭＳ ゴシック")
            etree.SubElement(rpr, q("ea")).set("typeface", "ＭＳ ゴシック")
            etree.SubElement(r, q("t")).text = "あああああ" + ch + "あいう"
    f = OUT / f"kinsoku_ea{ea_ln_brk or 'def'}.pptx"
    prs.save(f)
    return f


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    files = [build(None), build("0")]
    import pymupdf
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        for f in files:
            pres = app.Presentations.Open(str(f), WithWindow=False)
            pres.SaveAs(str(f.with_suffix(".pdf")), 32)
            pres.Close()
    finally:
        app.Quit()
    for f in files:
        for page in pymupdf.open(f.with_suffix(".pdf")):
            for b in page.get_text("dict")["blocks"]:
                lines = ["".join(s["text"] for s in l["spans"]) for l in b.get("lines", [])]
                if lines:
                    print(f.stem, " | ".join(lines))


if __name__ == "__main__":
    main()
