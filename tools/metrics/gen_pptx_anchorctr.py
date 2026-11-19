# -*- coding: utf-8 -*-
"""Probe: where does PowerPoint put text when bodyPr has anchorCtr="1"?

newdev 54542 s2's body box ("○" bullets, marL/indent, wrapped 14pt lines) is
the only anchorCtr box in the corpora, and PowerPoint draws its text about
2pt right of where a plain left-anchored box would. Each arm is one box; the
PDF origin x of every text line is printed, relative to the box's inner left.
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(__file__).resolve().parents[2] / "pipeline_data" / "pptx_probes" / "anchorctr"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
BOX_W = 300.0
LONG = "The quick brown fox jumps over the lazy dog again and again until it wraps"

# name: (paragraphs [(text, algn, marL_pt, indent_pt, bullet)], wrap, anchorCtr)
ARMS = {
    "short_ctr": ([("Short", "l", 0, 0, None)], "square", "1"),
    "short_noctr": ([("Short", "l", 0, 0, None)], "square", "0"),
    "two_lines_ctr": ([("Short", "l", 0, 0, None), ("A much longer second line", "l", 0, 0, None)], "square", "1"),
    "bullet_ctr": ([("Short bullet", "l", 30, -20, "○"), ("A much longer bullet line", "l", 30, -20, "○")], "square", "1"),
    "right_ctr": ([("Short", "r", 0, 0, None), ("A much longer second line", "r", 0, 0, None)], "square", "1"),
    "center_ctr": ([("Short", "ctr", 0, 0, None), ("A much longer second line", "ctr", 0, 0, None)], "square", "1"),
    "wrap_ctr": ([(LONG, "l", 0, 0, None)], "square", "1"),
    "nowrap_ctr": ([("Short", "l", 0, 0, None), ("A much longer second line", "l", 0, 0, None)], "none", "1"),
    "marl_ctr": ([("Short", "l", 40, 0, None), ("A much longer second line", "l", 0, 0, None)], "square", "1"),
}


def q(t: str) -> str:
    return f"{{{A}}}{t}"


def emu(pt: float) -> Emu:
    return Emu(int(round(pt * 12700)))


def build() -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = emu(960), emu(1080)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for k, (name, (paras, wrap, actr)) in enumerate(ARMS.items()):
        x, y = 20 + (k % 3) * 310, 20 + (k // 3) * 340
        box = slide.shapes.add_textbox(emu(x), emu(y), emu(BOX_W), emu(320))
        body = box.text_frame._txBody
        bp = body.find(q("bodyPr"))
        for a in ("lIns", "rIns", "tIns", "bIns"):
            bp.set(a, "0")
        bp.set("wrap", wrap)
        bp.set("anchor", "t")
        bp.set("anchorCtr", actr)
        for pel in body.findall(q("p")):
            body.remove(pel)
        for text, algn, marl, ind, bu in paras:
            p = etree.SubElement(body, q("p"))
            ppr = etree.SubElement(p, q("pPr"))
            ppr.set("algn", algn)
            ppr.set("marL", str(int(marl * 12700)))
            ppr.set("indent", str(int(ind * 12700)))
            if bu:
                etree.SubElement(ppr, q("buFont")).set("typeface", "Arial")
                etree.SubElement(ppr, q("buChar")).set("char", bu)
            else:
                etree.SubElement(ppr, q("buNone"))
            r = etree.SubElement(p, q("r"))
            rpr = etree.SubElement(r, q("rPr"))
            rpr.set("lang", "en-US")
            rpr.set("sz", "1800")
            etree.SubElement(rpr, q("latin")).set("typeface", "Arial")
            etree.SubElement(r, q("t")).text = text
    dst = OUT / "anchorctr.pptx"
    prs.save(dst)
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    f = build()
    import subprocess
    import time

    import pymupdf
    import win32com.client
    subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)
    time.sleep(1.5)
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(f), WithWindow=False)
        pdf = f.with_suffix(".pdf")
        pres.SaveAs(str(pdf), 32)
        pres.Close()
    finally:
        app.Quit()
    page = pymupdf.open(pdf)[0]
    spans = []
    for b in page.get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            for sp in line["spans"]:
                if sp["text"].strip():
                    spans.append((sp["origin"][0], sp["origin"][1], sp["bbox"][2], sp["text"].strip()))
    for k, name in enumerate(ARMS):
        x, y = 20 + (k % 3) * 310, 20 + (k // 3) * 340
        rows = sorted((round(oy - y, 1), round(ox - x, 2), round(x1 - x, 2), t[:14])
                      for ox, oy, x1, t in spans if x - 1 <= ox <= x + BOX_W + 200 and y <= oy <= y + 320)
        print(name.ljust(14), rows)


if __name__ == "__main__":
    main()
