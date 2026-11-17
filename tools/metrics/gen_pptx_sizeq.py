# -*- coding: utf-8 -*-
"""Probe: does PowerPoint lay text out at a font size rounded to 1/600 inch?

919f's PDF states every Meiryo UI run at a 0.12pt multiple (14 -> 14.04,
11 -> 11.04, 16 -> 15.96) and its full-width advances are 14.04, which is what
breaks "...講演/提|言" one character earlier than the 1/8pt master-unit advance
(14.0) would. Each arm is one zero-inset box holding N copies of one glyph, its
width set a hair off N x (nominal advance) so the two models disagree on
whether the run fits on one line. Output per arm: the line count in
PowerPoint's PDF and the line count COM reports (TextRange.Lines).
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\sizeq").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
N = 20
# (face, glyph, em advance, size pt, slack pt): box width = N*em*size + slack
ARMS = []
for face, ch, em in [("Meiryo UI", "漢", 1.0), ("ＭＳ ゴシック", "漢", 1.0), ("Arial", "n", 0.55615)]:
    for size in (14.0, 11.0, 16.0, 13.0, 12.0, 10.5):
        for slack in (-0.5, 0.5):
            ARMS.append((face, ch, em, size, slack))


def q(t: str) -> str:
    return f"{{{A}}}{t}"


def emu(pt: float) -> Emu:
    return Emu(int(round(pt * 12700)))


def build() -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = emu(960), emu(540)
    per = 24
    slides = [prs.slides.add_slide(prs.slide_layouts[6]) for _ in range((len(ARMS) + per - 1) // per)]
    for i, (face, ch, em, size, slack) in enumerate(ARMS):
        slide = slides[i // per]
        k = i % per
        w = N * em * size + slack
        box = slide.shapes.add_textbox(emu(10 + (k % 2) * 470), emu(10 + (k // 2) * 43), emu(w), emu(40))
        body = box.text_frame._txBody
        bp = body.find(q("bodyPr"))
        for a in ("lIns", "rIns", "tIns", "bIns"):
            bp.set(a, "0")
        bp.set("wrap", "square")
        for pel in body.findall(q("p")):
            body.remove(pel)
        p = etree.SubElement(body, q("p"))
        run = etree.SubElement(p, q("r"))
        rpr = etree.SubElement(run, q("rPr"))
        rpr.set("lang", "ja-JP")
        rpr.set("sz", str(int(size * 100)))
        etree.SubElement(rpr, q("latin")).set("typeface", face)
        etree.SubElement(rpr, q("ea")).set("typeface", face)
        etree.SubElement(run, q("t")).text = ch * N
    dst = OUT / "sizeq.pptx"
    prs.save(dst)
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    f = build()
    import pymupdf
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    com_lines = []
    try:
        pres = app.Presentations.Open(str(f), WithWindow=False)
        for s in pres.Slides:
            for sh in s.Shapes:
                com_lines.append(sh.TextFrame.TextRange.Lines().Count if False else sh.TextFrame2.TextRange.Lines.Count)
        pdf = f.with_suffix(".pdf")
        pres.SaveAs(str(pdf), 32)
        pres.Close()
    finally:
        app.Quit()
    doc = pymupdf.open(pdf)
    print("face".ljust(14), "size", "slack", "pdf_lines", "com_lines", "pdf_size")
    for i, (face, ch, em, size, slack) in enumerate(ARMS):
        page = doc[i // 24]
        k = i % 24
        x0, y0 = 10 + (k % 2) * 470, 10 + (k // 2) * 43
        ys, sz = set(), None
        for b in page.get_text("dict")["blocks"]:
            for line in b.get("lines", []):
                for sp in line["spans"]:
                    ox, oy = sp["origin"]
                    if x0 - 1 <= ox <= x0 + 460 and y0 - 2 <= oy <= y0 + 41 and sp["text"].strip():
                        ys.add(round(oy, 1))
                        sz = sp["size"]
        print(face.ljust(14), f"{size:5.1f}", f"{slack:+.1f}", len(ys), com_lines[i], sz)


if __name__ == "__main__":
    main()
