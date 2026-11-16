# -*- coding: utf-8 -*-
"""Probe: how PowerPoint lays out `a:bodyPr vert="eaVert"` text.

68ffbed2 slide 1 draws its flow-chart columns as eaVert text boxes; the
renderer has no vertical text, so each character wraps onto its own
horizontal line (pitch 1.2 em instead of PowerPoint's 1 em). Each arm is one
text box; COM reports every character's BoundLeft/Top/Width/Height relative
to the box, plus the box size after autofit.

Arms: plain column (anchor t, algn l), algn ctr, algn r, two paragraphs,
a wrapping paragraph, mixed Latin/kana/chōon, one inset at a time, anchor ctr.
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\eavert").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# label, (w, h) pt, paragraphs, algn, anchor, insets (l, t, r, b) pt or None
ARMS = [
    ("plain", (60, 240), ["あいうえお"], None, None, None),
    ("ctr", (60, 240), ["あいうえお"], "ctr", None, None),
    ("right", (60, 240), ["あいうえお"], "r", None, None),
    ("two", (90, 240), ["あいう", "かきく"], None, None, None),
    ("wrap", (90, 120), ["あいうえおかきくけこ"], None, None, None),
    ("mixed", (60, 300), ["AB12あーい、う。"], None, None, None),
    ("insL", (60, 240), ["あいう"], None, None, (36, 0, 0, 0)),
    ("insT", (60, 240), ["あいう"], None, None, (0, 36, 0, 0)),
    ("insR", (60, 240), ["あいう"], None, None, (0, 0, 36, 0)),
    ("ancC", (90, 240), ["あいう"], None, "ctr", None),
]
FS = 20
FONT = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("-") else "メイリオ"
TAG = {"メイリオ": "", "ＭＳ ゴシック": "_msg", "游ゴシック": "_yu"}.get(FONT, "_x")


def q(t: str) -> str:
    return f"{{{A}}}{t}"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    x = 20
    for label, (w, h), paras, algn, anchor, ins in ARMS:
        box = slide.shapes.add_textbox(Emu(int(Pt(x))), Emu(int(Pt(40))), Emu(int(Pt(w))), Emu(int(Pt(h))))
        x += w + 12
        body = box.text_frame._txBody
        bp = body.find(q("bodyPr"))
        bp.set("vert", "eaVert")
        bp.set("wrap", "square")
        if anchor:
            bp.set("anchor", anchor)
        if ins:
            for k, v in zip(("lIns", "tIns", "rIns", "bIns"), ins):
                bp.set(k, str(int(Pt(v))))
        for e in list(bp):
            bp.remove(e)  # no autofit
        for pel in body.findall(q("p")):
            body.remove(pel)
        for text in paras:
            p = etree.SubElement(body, q("p"))
            if algn:
                etree.SubElement(p, q("pPr")).set("algn", algn)
            r = etree.SubElement(p, q("r"))
            rpr = etree.SubElement(r, q("rPr"))
            rpr.set("lang", "ja-JP")
            rpr.set("sz", str(FS * 100))
            etree.SubElement(rpr, q("latin")).set("typeface", FONT)
            etree.SubElement(rpr, q("ea")).set("typeface", FONT)
            etree.SubElement(r, q("t")).text = text
    dst = OUT / f"eavert{TAG}.pptx"
    prs.save(dst)
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(dst), WithWindow=False)
        for k, (label, *_r) in enumerate(ARMS, start=1):
            sh = pres.Slides(1).Shapes(k)
            tr = sh.TextFrame.TextRange
            cells = []
            for i in range(1, tr.Length + 1):
                c = tr.Characters(i, 1)
                if c.Text in ("\r", "\n"):
                    continue
                cells.append(f"{c.Text}({c.BoundLeft - sh.Left:.1f},{c.BoundTop - sh.Top:.1f},"
                             f"{c.BoundWidth:.1f}x{c.BoundHeight:.1f})")
            print(f"{label:6s} box {sh.Width:.1f}x{sh.Height:.1f}: " + " ".join(cells))
        pres.SaveAs(str(dst.with_suffix(".pdf")), 32)
        pres.Close()
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
