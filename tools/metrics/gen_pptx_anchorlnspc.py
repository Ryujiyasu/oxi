# -*- coding: utf-8 -*-
"""Probe: first and second baselines of a two-paragraph box, per family x
(anchor, lnSpc).

54542be9's centre-anchored lnSpc-130% boxes sit 0.7-1.2pt LOWER in PowerPoint
than in Oxi for MS PGothic / Noto Sans JP while Yu Gothic matches; top-anchored
single lines match for all three. Grid: rows = families, columns = configs,
each a 60pt-tall box with default insets holding two 11pt paragraphs.
Output: per cell, baseline1 - box top and baseline2 - baseline1, from
PowerPoint's own PDF.
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\anchorlnspc").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
FAMS = ["ＭＳ Ｐゴシック", "Noto Sans JP", "Yu Gothic", "Calibri", "Arial", "Meiryo UI"]
CONFIGS = [("t", None, 0), ("t", 130000, 0), ("ctr", None, 0), ("ctr", 130000, 0), ("b", 130000, 0), ("t", 80000, 0),
           ("b", 80000, 0), ("ctr", 80000, 0), ("b", 150000, 0), ("t", 150000, 0), ("b", 130000, 1200), ("b", None, 1200)]
PER_SLIDE = 6
X0, Y0, DX, DY, W, H = 20, 20, 115, 85, 108, 60  # pt
# Mode (argv[1]): "flp" sets bodyPr spcFirstLastPara="1" (Google Slides exports
# carry it); "over" shrinks the box to 15pt so the two lines overflow it.
MODE = sys.argv[1] if len(sys.argv) > 1 else ""
if "over" in MODE:
    H = 15
if MODE:
    OUT = OUT.parent / f"anchorlnspc_{MODE}"


def q(t: str) -> str:
    return f"{{{A}}}{t}"


def emu(pt: float) -> Emu:
    return Emu(int(pt * 12700))


def build() -> Path:
    prs = Presentation()
    slides = [prs.slides.add_slide(prs.slide_layouts[6]) for _ in range((len(CONFIGS) + PER_SLIDE - 1) // PER_SLIDE)]
    for r, fam in enumerate(FAMS):
        for c, (anchor, pct, aft) in enumerate(CONFIGS):
            slide = slides[c // PER_SLIDE]
            box = slide.shapes.add_textbox(emu(X0 + (c % PER_SLIDE) * DX), emu(Y0 + r * DY), emu(W), emu(H))
            body = box.text_frame._txBody
            bp = body.find(q("bodyPr"))
            bp.set("anchor", anchor)
            bp.set("wrap", "none")
            if "flp" in MODE:
                bp.set("spcFirstLastPara", "1")
            for pel in body.findall(q("p")):
                body.remove(pel)
            for j in range(2):
                p = etree.SubElement(body, q("p"))
                if pct or aft:
                    ppr = etree.SubElement(p, q("pPr"))
                    if pct:
                        etree.SubElement(etree.SubElement(ppr, q("lnSpc")), q("spcPct")).set("val", str(pct))
                    if aft:
                        etree.SubElement(etree.SubElement(ppr, q("spcAft")), q("spcPts")).set("val", str(aft))
                run = etree.SubElement(p, q("r"))
                rpr = etree.SubElement(run, q("rPr"))
                rpr.set("lang", "ja-JP")
                rpr.set("sz", "1100")
                etree.SubElement(rpr, q("latin")).set("typeface", fam)
                etree.SubElement(rpr, q("ea")).set("typeface", fam)
                etree.SubElement(run, q("t")).text = f"{r}{c:x}{j}Ab漢"
    dst = OUT / "anchorlnspc.pptx"
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
    base = {}
    for page in pymupdf.open(pdf):
      for b in page.get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            sp = line["spans"][0]
            t = sp["text"]
            if len(t) < 3 or not (t[0].isdigit() and t[2].isdigit()):
                continue
            r, c, j = int(t[0]), int(t[1], 16), int(t[2])
            base[(r, c, j)] = sp["origin"][1] - (Y0 + r * DY)
    print("family".ljust(16), "  ".join(f"{a}/{(p or 100000)//1000}/{f//100}".rjust(13) for a, p, f in CONFIGS))
    for r, fam in enumerate(FAMS):
        cells = []
        for c in range(len(CONFIGS)):
            b1, b2 = base.get((r, c, 0)), base.get((r, c, 1))
            cells.append(f"{b1:6.2f}+{b2 - b1:5.2f}" if b1 is not None and b2 is not None else "?".rjust(13))
        print(fam.ljust(16), "  ".join(cells))


if __name__ == "__main__":
    main()
