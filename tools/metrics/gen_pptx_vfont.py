# -*- coding: utf-8 -*-
"""Probe: does PowerPoint draw text in an installed VARIABLE font, and where is
the first baseline?

54542be9's runs ask for "Noto Sans JP"; this machine has NotoSansJP-VF.ttf
installed and GDI serves the name, yet PowerPoint's own PDF draws the text in
MS PGothic. Each arm is a 20pt text box (tIns 0, top anchored) holding
"Abc あいう漢字" in one family set as latin AND ea; the PDF span fonts and the
first baseline (origin y - box top) say which face drew it and which face's
ascent split placed the line.
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\vfont").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
FAMS = ["Noto Sans JP", "ＭＳ Ｐゴシック", "Bahnschrift", "Sitka Text", "Yu Gothic", "Meiryo UI"]
# Second slide: the same families in a CENTRE-anchored 60pt box holding two
# 11pt paragraphs at lnSpc 130% (54542's body boxes).
CTR_FAMS = ["Noto Sans JP", "ＭＳ Ｐゴシック", "Yu Gothic"]
TOP0, STEP = 457200, 900000


def q(t: str) -> str:
    return f"{{{A}}}{t}"


def build() -> Path:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for k, fam in enumerate(FAMS):
        box = slide.shapes.add_textbox(Emu(457200), Emu(TOP0 + k * STEP), Emu(7000000), Emu(800000))
        body = box.text_frame._txBody
        bp = body.find(q("bodyPr"))
        for a in ("tIns", "bIns", "lIns"):
            bp.set(a, "0")
        bp.set("anchor", "t")
        for pel in body.findall(q("p")):
            body.remove(pel)
        p = etree.SubElement(body, q("p"))
        r = etree.SubElement(p, q("r"))
        rpr = etree.SubElement(r, q("rPr"))
        rpr.set("lang", "ja-JP")
        rpr.set("sz", "2000")
        etree.SubElement(rpr, q("latin")).set("typeface", fam)
        etree.SubElement(rpr, q("ea")).set("typeface", fam)
        etree.SubElement(r, q("t")).text = f"{k} Abc あいう漢字"
    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    for k, fam in enumerate(CTR_FAMS):
        box = s2.shapes.add_textbox(Emu(457200), Emu(TOP0 + k * STEP * 2), Emu(7000000), Emu(762000))
        body = box.text_frame._txBody
        bp = body.find(q("bodyPr"))
        bp.set("anchor", "ctr")
        for pel in body.findall(q("p")):
            body.remove(pel)
        for j in range(2):
            p = etree.SubElement(body, q("p"))
            ppr = etree.SubElement(p, q("pPr"))
            etree.SubElement(etree.SubElement(ppr, q("lnSpc")), q("spcPct")).set("val", "130000")
            r = etree.SubElement(p, q("r"))
            rpr = etree.SubElement(r, q("rPr"))
            rpr.set("lang", "ja-JP")
            rpr.set("sz", "1100")
            etree.SubElement(rpr, q("latin")).set("typeface", fam)
            etree.SubElement(rpr, q("ea")).set("typeface", fam)
            etree.SubElement(r, q("t")).text = f"{k}{j} Abc あいう漢字"
    dst = OUT / "vfont.pptx"
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
    doc = pymupdf.open(pdf)
    for x in doc[0].get_fonts(full=True):
        b = doc.extract_font(x[0])
        try:
            from fontTools.ttLib import TTFont
            import io
            real = TTFont(io.BytesIO(b[3]))["name"].getDebugName(4)
        except Exception:
            real = "?"
        print("font", x[3], "->", real)
    for b in doc[0].get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            for sp in line["spans"]:
                y = sp["origin"][1]
                k = min(range(len(FAMS)), key=lambda i: abs(y - (TOP0 + i * STEP) / 12700 - 18))
                top = (TOP0 + k * STEP) / 12700
                print(FAMS[k], "baseline-top=%.2f" % (y - top), f"{sp['font']}:{sp['text']}")
    for b in doc[1].get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            sp = line["spans"][0]
            k = int(sp["text"][0])
            top = (TOP0 + k * STEP * 2) / 12700
            print("ctr", CTR_FAMS[k], sp["text"][:2], "baseline-top=%.2f" % (sp["origin"][1] - top))

if __name__ == "__main__":
    main()
