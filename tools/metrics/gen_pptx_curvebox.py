# -*- coding: utf-8 -*-
"""Probe: which rectangle does PowerPoint stretch a shape's blipFill into when
the shape's custom geometry does not reach the edges of its box?

d50 s9 (dev), 46 s18 and 31 s2 (blind) fill Bezier-outlined custGeom shapes
with photos; PowerPoint's picture lands on the bounding box of the DRAWN
outline, not on the shape's xfrm box. Each arm is one custGeom shape filled
with the same test image; the PDF keeps every placed image's full rectangle
(even where the outline clips it), so it is read back directly and compared
with the two predictions.
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(__file__).resolve().parents[2] / "pipeline_data" / "pptx_probes" / "curvebox"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
BOX = 220.0  # pt, every arm's box is BOX x BOX
PW = PH = 1000  # path coordinate space

# A blob whose cubic control points touch the box edges but whose curve does
# not: x spans 125..875, y about 47..953 of the 1000-unit path.
BLOB = [("moveTo", [(500, 50)]),
        ("cubicBezTo", [(1000, 0), (1000, 1000), (500, 950)]),
        ("cubicBezTo", [(0, 1000), (0, 0), (500, 50)]),
        ("close", [])]
# Straight segments that stop short of the right and bottom edges.
POLY = [("moveTo", [(0, 0)]), ("lnTo", [(800, 100)]), ("lnTo", [(600, 700)]),
        ("lnTo", [(100, 800)]), ("close", [])]
# Quadratic segments: control points on the edges, curve inside them.
QUAD = [("moveTo", [(500, 100)]),
        ("quadBezTo", [(1000, 500), (500, 900)]),
        ("quadBezTo", [(0, 500), (500, 100)]),
        ("close", [])]
RECT = [("moveTo", [(0, 0)]), ("lnTo", [(1000, 0)]), ("lnTo", [(1000, 1000)]),
        ("lnTo", [(0, 1000)]), ("close", [])]

# name: (path, fillRect l, t, r, b in percent)
ARMS = {
    "blob": (BLOB, (0, 0, 0, 0)),
    "blob_negfill": (BLOB, (-20, -5, -10, 5)),
    "poly": (POLY, (0, 0, 0, 0)),
    "quad": (QUAD, (0, 0, 0, 0)),
    "rect_control": (RECT, (0, 0, 0, 0)),
    "poly_negfill": (POLY, (-10, -10, -10, -10)),
}


def q(ns: str, t: str) -> str:
    return f"{{{ns}}}{t}"


def emu(pt: float) -> int:
    return int(round(pt * 12700))


def pos(k: int) -> tuple[float, float]:
    return 40 + (k % 3) * 300, 40 + (k // 3) * 300


def test_image() -> bytes:
    """A 4x4 grid of distinct colours with a black frame."""
    im = Image.new("RGB", (400, 400), "white")
    d = ImageDraw.Draw(im)
    for i in range(4):
        for j in range(4):
            d.rectangle([j * 100, i * 100, j * 100 + 99, i * 100 + 99],
                        fill=((j * 70 + 30) % 256, (i * 60 + 40) % 256, ((i + j) * 40 + 90) % 256))
    d.rectangle([0, 0, 399, 399], outline="black", width=6)
    buf = io.BytesIO()
    im.save(buf, "PNG")
    return buf.getvalue()


def build() -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(emu(960)), Emu(emu(660))
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # Register the image once through a throwaway picture to get its rId.
    pic = slide.shapes.add_picture(io.BytesIO(test_image()), 0, 0)
    rid = pic._element.find(f".//{q(A, 'blip')}").get(q(R, "embed"))
    pic._element.getparent().remove(pic._element)
    tree = slide.shapes._spTree
    for k, (name, (path, (l, t, r, b))) in enumerate(ARMS.items()):
        x, y = pos(k)
        sp = etree.SubElement(tree, q(P, "sp"))
        nv = etree.SubElement(sp, q(P, "nvSpPr"))
        c = etree.SubElement(nv, q(P, "cNvPr"))
        c.set("id", str(100 + k))
        c.set("name", name)
        etree.SubElement(nv, q(P, "cNvSpPr"))
        etree.SubElement(nv, q(P, "nvPr"))
        sppr = etree.SubElement(sp, q(P, "spPr"))
        xf = etree.SubElement(sppr, q(A, "xfrm"))
        off = etree.SubElement(xf, q(A, "off"))
        off.set("x", str(emu(x)))
        off.set("y", str(emu(y)))
        ext = etree.SubElement(xf, q(A, "ext"))
        ext.set("cx", str(emu(BOX)))
        ext.set("cy", str(emu(BOX)))
        cg = etree.SubElement(sppr, q(A, "custGeom"))
        for tag in ("avLst", "gdLst", "ahLst", "cxnLst"):
            etree.SubElement(cg, q(A, tag))
        rect = etree.SubElement(cg, q(A, "rect"))
        for a_, v in (("l", "l"), ("t", "t"), ("r", "r"), ("b", "b")):
            rect.set(a_, v)
        pl = etree.SubElement(cg, q(A, "pathLst"))
        pe = etree.SubElement(pl, q(A, "path"))
        pe.set("w", str(PW))
        pe.set("h", str(PH))
        for cmd, pts in path:
            el = etree.SubElement(pe, q(A, cmd))
            for px, py in pts:
                pt = etree.SubElement(el, q(A, "pt"))
                pt.set("x", str(px))
                pt.set("y", str(py))
        bf = etree.SubElement(sppr, q(A, "blipFill"))
        bf.set("rotWithShape", "1")
        etree.SubElement(bf, q(A, "blip")).set(q(R, "embed"), rid)
        st = etree.SubElement(bf, q(A, "stretch"))
        fr = etree.SubElement(st, q(A, "fillRect"))
        for a_, v in (("l", l), ("t", t), ("r", r), ("b", b)):
            fr.set(a_, str(int(v * 1000)))
        etree.SubElement(etree.SubElement(sppr, q(A, "ln")), q(A, "noFill"))
    dst = OUT / "curvebox.pptx"
    prs.save(dst)
    return dst


def curve_bbox(path: list) -> tuple[float, float, float, float]:
    """Bounding box of the drawn outline, in path units."""
    xs, ys, cur = [], [], None
    for cmd, pts in path:
        if cmd in ("moveTo", "lnTo"):
            cur = pts[0]
            xs.append(cur[0])
            ys.append(cur[1])
        elif cmd in ("cubicBezTo", "quadBezTo"):
            ctrl = [cur] + pts
            n = len(ctrl) - 1
            for i in range(401):
                u = i / 400
                bx = by = 0.0
                for j, (cx, cy) in enumerate(ctrl):
                    from math import comb
                    w = comb(n, j) * (1 - u) ** (n - j) * u ** j
                    bx += w * cx
                    by += w * cy
                xs.append(bx)
                ys.append(by)
            cur = pts[-1]
    return min(xs), min(ys), max(xs), max(ys)


def predict(x: float, y: float, bbox: tuple, fill: tuple) -> list[float]:
    l, t, r, b = (v / 100 for v in fill)
    bx0, by0 = x + bbox[0] / PW * BOX, y + bbox[1] / PH * BOX
    bw, bh = (bbox[2] - bbox[0]) / PW * BOX, (bbox[3] - bbox[1]) / PH * BOX
    return [round(v, 1) for v in (bx0 + l * bw, by0 + t * bh, bx0 + bw * (1 - r), by0 + bh * (1 - b))]


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
    placed = [i["bbox"] for i in pymupdf.open(pdf)[0].get_image_info()]
    for k, (name, (path, fill)) in enumerate(ARMS.items()):
        x, y = pos(k)
        box = predict(x, y, (0, 0, PW, PH), fill)
        curve = predict(x, y, curve_bbox(path), fill)
        near = [b for b in placed if abs(b[0] - curve[0]) < 80 and abs(b[1] - curve[1]) < 80]
        got = [round(v, 1) for v in near[0]] if near else None
        print(f"{name:14s} box {box}  curve {curve}  PowerPoint {got}")


if __name__ == "__main__":
    main()
