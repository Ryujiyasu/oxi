# -*- coding: utf-8 -*-
"""Probe: where does a tab in a PowerPoint paragraph land?

252cbc8c slide 2 / 640d6e4e / abde73e7 line up columns with tabs; the pptx
renderer has no tab stops at all (a tab is measured as a glyph). Each
paragraph below is one arm; COM reads BoundLeft of the character after the
LAST tab (relative to the shape's left edge, lIns=0) and of its end.

Arms: defTabSz multiples (where are they measured from?), marL / hanging
indent, a custom tabLst stop and what follows it, algn r / ctr, two tabs in a
row, a full-width (Japanese) run.
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\tabs").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
IN = 914400

# (label, marL, indent, defTabSz, tabs[(pos, algn)], text)
ARMS = [
    ("d1", 0, 0, None, [], "A\tB"),
    ("d_long", 0, 0, None, [], "AAAAAAAAAAAAAAAAAAAA\tB"),
    ("d_two", 0, 0, None, [], "A\t\tB"),
    ("marL1", IN, 0, None, [], "A\tB"),
    ("marL1_long", IN, 0, None, [], "AAAAAAAAAAAA\tB"),
    ("hang", IN, -IN, None, [], "A\tB"),
    ("hang_long", IN, -IN // 2, None, [], "AAAAAAAAAAAA\tB"),
    ("dts5", 0, 0, 500000, [], "A\tB"),
    ("dts5_two", 0, 0, 500000, [], "A\t\tB"),
    ("dts5_marL", 700000, 0, 500000, [], "A\tB"),
    ("cust", 0, 0, None, [(300000, "l")], "A\tB"),
    ("cust_past", 0, 0, None, [(300000, "l")], "AAAAAAAAA\tB"),
    ("cust_two", 0, 0, None, [(300000, "l")], "A\t\tB"),
    ("cust_marL", IN, 0, None, [(300000, "l")], "A\tB"),
    ("cust_far", 0, 0, None, [(3000000, "l")], "A\tB"),
    ("cust_far2", 0, 0, None, [(3000000, "l")], "A\t\tB"),
    ("hang_cust", IN, -IN, None, [(300000, "l")], "A\t\tB"),
    ("right", 0, 0, None, [(3000000, "r")], "A\tBBBB"),
    ("ctr", 0, 0, None, [(3000000, "ctr")], "A\tBBBB"),
    ("ja", 0, 0, 972616, [], "ステップ１：担当チームを決めよう\t\t　（４）"),
    ("ja1", 0, 0, 972616, [], "ステップ３：公開データの準備をしよう\t　（３）"),
]


def q(tag: str) -> str:
    return f"{{{A}}}{tag}"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Emu(IN // 2), Emu(IN // 4), Emu(11 * IN), Emu(6 * IN))
    body = box.text_frame._txBody
    bp = body.find(q("bodyPr"))
    for k in ("lIns", "rIns", "tIns", "bIns"):
        bp.set(k, "0")
    bp.set("wrap", "none")
    for pel in body.findall(q("p")):
        body.remove(pel)
    for label, marl, ind, dts, tabs, text in ARMS:
        p = etree.SubElement(body, q("p"))
        ppr = etree.SubElement(p, q("pPr"))
        ppr.set("marL", str(marl))
        ppr.set("indent", str(ind))
        if dts:
            ppr.set("defTabSz", str(dts))
        if tabs:
            tl = etree.SubElement(ppr, q("tabLst"))
            for pos, al in tabs:
                t = etree.SubElement(tl, q("tab"))
                t.set("pos", str(pos))
                t.set("algn", al)
        etree.SubElement(ppr, q("buNone"))
        r = etree.SubElement(p, q("r"))
        rpr = etree.SubElement(r, q("rPr"))
        rpr.set("lang", "ja-JP" if label.startswith("ja") else "en-US")
        rpr.set("sz", "1400")
        etree.SubElement(rpr, q("latin")).set("typeface", "Arial")
        etree.SubElement(rpr, q("ea")).set("typeface", "ＭＳ ゴシック")
        etree.SubElement(r, q("t")).text = text
    dst = OUT / "tabs.pptx"
    prs.save(dst)
    if "--no-com" in sys.argv:
        return
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(dst), WithWindow=False)
        sh = pres.Slides(1).Shapes(1)
        left = sh.Left
        tr = sh.TextFrame.TextRange
        for i, (label, *_rest, text) in enumerate(ARMS, start=1):
            para = tr.Paragraphs(i)
            k = text.rfind("\t") + 1  # 0-based index of the char after the last tab
            after = para.Characters(k + 1, 1)
            last = para.Characters(len(text), 1)
            first = para.Characters(1, 1)
            print(f"{label:11s} first={first.BoundLeft - left:8.2f} after_tab={after.BoundLeft - left:8.2f}"
                  f" end={last.BoundLeft + last.BoundWidth - left:8.2f}")
        pres.SaveAs(str(dst.with_suffix(".pdf")), 32)
        pres.Close()
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
