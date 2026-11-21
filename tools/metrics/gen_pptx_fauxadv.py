# -*- coding: utf-8 -*-
"""Probe: at what advances does PowerPoint set a bold run whose family has no
bold face (so the bold is synthesised), and how does that depend on size?

d31 s16 (dev): Paytone One, regular-only, b=1 at 77.18pt. PowerPoint's pen
advances are 0.984 x design; GDI's synthesised bold measured at a 16384px
probe (what Oxi asks) gives 0.980, at 512px 0.984, at 128px 0.997 -- GDI's
synthetic bold is about 0.980 x design + ~1-2px per glyph. Oxi's 0.4% narrow
line kept "involves decorating" on one line where PowerPoint broke it.

Each arm is one unwrapped text box: one face, one size, bold or not. The PDF
keeps every glyph's origin, so the pen advance of each glyph is read back and
divided by the face's design advance (from the installed font file).

    python tools/metrics/gen_pptx_fauxadv.py           # build + PowerPoint PDF + table
    python tools/metrics/gen_pptx_fauxadv.py "Arial Black=ariblk.ttf" ... --sizes 36,77.18,150
"""
from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(__file__).resolve().parents[2] / "pipeline_data" / "pptx_probes" / "fauxadv"
FACES = {"Impact": "impact.ttf", "Calibri Light": "calibril.ttf"}
SIZES = [10, 14, 18, 24, 36, 48, 60, 77.18, 100, 150]
TEXT = "Handgloves quick brown fox"


def build() -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(int(2400 * 12700)), Emu(int(1400 * 12700))
    for face in FACES:
        for bold in (True, False):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            y = 20.0
            for size in SIZES:
                tb = slide.shapes.add_textbox(Pt(20), Pt(y), Pt(2300), Pt(size * 1.4))
                tf = tb.text_frame
                tf.word_wrap = False
                body = tf._txBody.find("{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr")
                for k in ("lIns", "tIns", "rIns", "bIns"):
                    body.set(k, "0")
                r = tf.paragraphs[0].add_run()
                r.text = TEXT
                r.font.size = Pt(size)
                r.font.bold = bold
                r.font.name = face
                y += size * 1.4 + 6
    dst = OUT / "fauxadv.pptx"
    prs.save(dst)
    return dst


def export(src: Path) -> Path:
    import win32com.client
    subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)
    time.sleep(3)
    app = win32com.client.DispatchEx("PowerPoint.Application")
    pdf = src.with_suffix(".pdf")
    try:
        p = app.Presentations.Open(os.path.normpath(src), ReadOnly=True, WithWindow=False)
        p.SaveAs(os.path.normpath(pdf), 32)
        p.Close()
    finally:
        app.Quit()
    return pdf


def read(pdf: Path) -> None:
    import pymupdf
    from fontTools.ttLib import TTFont
    fonts = {}
    for face, file in FACES.items():
        f = TTFont(os.path.join(os.environ["WINDIR"], "Fonts", file))
        fonts[face] = (f["hmtx"], f.getBestCmap(), f["head"].unitsPerEm)
    doc = pymupdf.open(pdf)
    k = 0
    for face in FACES:
        hm, cmap, upm = fonts[face]
        for bold in (True, False):
            page = doc[k]
            k += 1
            rows = []
            for b in page.get_text("rawdict")["blocks"]:
                for line in b.get("lines", []):
                    for s in line["spans"]:
                        ch = s["chars"]
                        pen = ch[-1]["origin"][0] - ch[0]["origin"][0]
                        des = sum(hm[cmap[ord(c["c"])]][0] for c in ch[:-1]) / upm * s["size"]
                        rows.append((s["size"], pen / des, s["font"]))
            for size, ratio, fname in sorted(rows):
                print(f"{face:14s} {'bold' if bold else 'reg ':4s} {size:7.2f}pt  pen/design {ratio:.4f}  {fname}")


def main() -> None:
    global FACES, SIZES
    args = sys.argv[1:]
    if "--sizes" in args:
        i = args.index("--sizes")
        SIZES = [float(x) for x in args[i + 1].split(",")]
        del args[i : i + 2]
    if args:
        FACES = dict(a.split("=", 1) for a in args)
    OUT.mkdir(parents=True, exist_ok=True)
    read(export(build()))


if __name__ == "__main__":
    main()
