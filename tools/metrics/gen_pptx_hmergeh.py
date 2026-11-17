# -*- coding: utf-8 -*-
"""Probe: does a merged-away cell (hMerge / vMerge continuation) size its row?

252cbc8c s96's category tables open with a gridSpan=2 header whose hMerge
continuation holds only an empty paragraph marked sz=1200, while the visible
cell is 10pt. Each arm is a 2x2 table: row 1 = a 10pt "AB" cell spanning both
columns plus a continuation cell (hMerge) or, for the vMerge arms, row 1's
right cell spans two rows and row 2's right cell is the continuation. The
continuation's content varies (empty paragraph at 12/24/40pt, or 40pt text).
Output per arm: COM Rows(1..2).Height, the rendered row heights.
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\hmergeh").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
# (kind, hidden content: ("empty", sz) or ("text", sz))
ARMS = [(k, c) for k in ("h", "v") for c in (("empty", 12), ("empty", 24), ("empty", 40), ("text", 40))]
ARMS.append(("none", ("empty", 12)))


def q(t: str) -> str:
    return f"{{{A}}}{t}"


def emu(pt: float) -> Emu:
    return Emu(int(round(pt * 12700)))


def set_cell(tc, text: str | None, sz: int) -> None:
    body = tc.find(q("txBody"))
    for p in body.findall(q("p")):
        body.remove(p)
    p = etree.SubElement(body, q("p"))
    if text:
        r = etree.SubElement(p, q("r"))
        rpr = etree.SubElement(r, q("rPr"))
        rpr.set("lang", "en-US")
        rpr.set("sz", str(sz * 100))
        etree.SubElement(rpr, q("latin")).set("typeface", "Arial")
        etree.SubElement(r, q("t")).text = text
    else:
        e = etree.SubElement(p, q("endParaRPr"))
        e.set("lang", "en-US")
        e.set("sz", str(sz * 100))
        etree.SubElement(e, q("latin")).set("typeface", "Arial")


def build() -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = emu(960), emu(540)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for i, (kind, (ck, csz)) in enumerate(ARMS):
        x, y = 20 + (i % 5) * 185, 20 + (i // 5) * 200
        gf = slide.shapes.add_table(2, 2, emu(x), emu(y), emu(160), emu(20))
        tbl = gf.table
        tbl.rows[0].height = emu(8)
        tbl.rows[1].height = emu(8)
        t = tbl._tbl
        trs = t.findall(q("tr"))
        c00, c01 = trs[0].findall(q("tc"))
        c10, c11 = trs[1].findall(q("tc"))
        set_cell(c10, "x", 10)
        hidden_text = "WW" if ck == "text" else None
        if kind == "h":
            c00.set("gridSpan", "2")
            c01.set("hMerge", "1")
            set_cell(c00, "AB", 10)
            set_cell(c01, hidden_text, csz)
            set_cell(c11, "y", 10)
        elif kind == "v":
            c01.set("rowSpan", "2")
            c11.set("vMerge", "1")
            set_cell(c00, "AB", 10)
            set_cell(c01, "z", 10)
            set_cell(c11, hidden_text, csz)
        else:
            set_cell(c00, "AB", 10)
            set_cell(c01, hidden_text, csz)
            set_cell(c11, "y", 10)
    dst = OUT / "hmergeh.pptx"
    prs.save(dst)
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    f = build()
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(f), WithWindow=False)
        rows = []
        for sh in pres.Slides(1).Shapes:
            if sh.HasTable:
                t = sh.Table
                rows.append((t.Rows(1).Height, t.Rows(2).Height))
        pres.SaveAs(str(f.with_suffix(".pdf")), 32)
        pres.Close()
    finally:
        app.Quit()
    for (kind, c), (h1, h2) in zip(ARMS, rows):
        print(kind, c, f"row1={h1:.2f} row2={h2:.2f}")


if __name__ == "__main__":
    main()
