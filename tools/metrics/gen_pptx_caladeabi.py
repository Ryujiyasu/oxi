# -*- coding: utf-8 -*-
"""Probe: how does PowerPoint draw b/i combinations of a family whose
Bold Italic file names its subfamily "Italic"?

The system's Caladea-BoldItalic.ttf (weight 700, macStyle 3) carries name ID 2
"Italic", the same as Caladea-Italic.ttf. Blind 47 slide 2's b=1 i=1 runs come
out of PowerPoint looking like the plain italic. One box per (b, i) arm, no
embedded fonts; output the PDF span font, its embedded weight and the run's
width (a real bold is wider).
"""
from __future__ import annotations

import io
import os
import subprocess
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(__file__).resolve().parents[2] / "pipeline_data" / "pptx_probes" / "caladeabi"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
FACES = sys.argv[1:] or ["Caladea"]
ARMS = [(b, i) for b in (0, 1) for i in (0, 1)]


def q(t: str) -> str:
    return f"{{{A}}}{t}"


def emu(pt: float) -> Emu:
    return Emu(int(round(pt * 12700)))


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width, prs.slide_height = emu(960), emu(540)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    k = 0
    for face in FACES:
        for b, i in ARMS:
            box = slide.shapes.add_textbox(emu(20 + (k % 2) * 470), emu(20 + (k // 2) * 60), emu(460), emu(50))
            body = box.text_frame._txBody
            for pel in body.findall(q("p")):
                body.remove(pel)
            p = etree.SubElement(body, q("p"))
            r = etree.SubElement(p, q("r"))
            rpr = etree.SubElement(r, q("rPr"))
            rpr.set("lang", "en-US")
            rpr.set("sz", os.environ.get("PROBE_SZ", "2800"))
            rpr.set("b", str(b))
            rpr.set("i", str(i))
            etree.SubElement(rpr, q("latin")).set("typeface", face)
            etree.SubElement(r, q("t")).text = os.environ.get("PROBE_TEXT", f"b{b}i{i} Hamburgefonstiv")
            k += 1
    f = OUT / "caladeabi.pptx"
    prs.save(f)
    import pymupdf
    import win32com.client
    from fontTools.ttLib import TTFont
    subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(f), WithWindow=False)
        pres.SaveAs(str(f.with_suffix(".pdf")), 32)
        pres.Close()
    finally:
        app.Quit()
    doc = pymupdf.open(f.with_suffix(".pdf"))
    page = doc[0]
    info = {}
    for x in page.get_fonts():
        try:
            t = TTFont(io.BytesIO(doc.extract_font(x[0])[3]))
            info[x[3].split("+")[-1]] = (t["OS/2"].usWeightClass, t["name"].getDebugName(4))
        except Exception as e:  # noqa: BLE001
            info[x[3].split("+")[-1]] = str(e)[:30]
    for bl in page.get_text("dict")["blocks"]:
        for line in bl.get("lines", []):
            for sp in line["spans"]:
                if sp["text"].strip():
                    print(sp["text"][:5], sp["font"], info.get(sp["font"]), round(sp["bbox"][2] - sp["bbox"][0], 1))


if __name__ == "__main__":
    main()
