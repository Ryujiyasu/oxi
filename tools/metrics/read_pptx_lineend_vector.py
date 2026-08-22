# -*- coding: utf-8 -*-
"""Derive the line-end (arrowhead) size law from PowerPoint's own PDF vectors.

A rasterised arrowhead on a 0.75pt line is about a dozen pixels across, and at
that size antialiasing is worth more than the quantity being measured. The PDF
carries the head as a real filled PATH, so its extent can be read exactly --
the same trick `read_pptx_dash.py` uses to read a dash array verbatim.

For each decorated connector the slide XML declares, this finds the filled path
sitting at that endpoint and reports its size ACROSS and ALONG the line, in
points and in line widths, grouped by (type, w, len, line width).

Usage: python tools/metrics/read_pptx_lineend_vector.py [deck-prefix ...]
"""
from __future__ import annotations

import glob
import math
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path
import xml.etree.ElementTree as ET

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
DEV = Path(r"pipeline_data\pptx_benchmark\dev")


def path_points(d: dict) -> list[tuple[float, float]]:
    """Every point a pymupdf drawing touches."""
    out = []
    for item in d["items"]:
        for p in item[1:]:
            if isinstance(p, pymupdf.Point):
                out.append((p.x, p.y))
            elif isinstance(p, pymupdf.Rect):
                out += [(p.x0, p.y0), (p.x1, p.y1)]
            elif isinstance(p, pymupdf.Quad):
                out += [(q.x, q.y) for q in (p.ul, p.ur, p.ll, p.lr)]
    return out


def main() -> None:
    wanted = sys.argv[1:]
    rows = defaultdict(list)
    for deck_path in sorted(glob.glob(str(DEV / "pptx" / "*.pptx"))):
        deck = Path(deck_path).stem
        if wanted and not any(deck.startswith(w) for w in wanted):
            continue
        pdf_path = DEV / "pdf" / f"{deck}.pdf"
        if not pdf_path.exists():
            continue
        zf = zipfile.ZipFile(deck_path)
        pres = ET.fromstring(zf.read("ppt/presentation.xml"))
        sz = pres.find(f"{P}sldSz")
        sw_pt = int(sz.get("cx")) / 12700
        pdf = pymupdf.open(pdf_path)
        for name in zf.namelist():
            m = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
            if not m:
                continue
            sn = int(m.group(1))
            raw = zf.read(name)
            if b"End" not in raw or sn > pdf.page_count:
                continue
            page = pdf[sn - 1]
            k = page.rect.width / sw_pt          # PDF pt per slide pt
            drawings = None
            for sp in ET.fromstring(raw).iter(f"{P}cxnSp"):
                ln = sp.find(f".//{A}ln")
                xf = sp.find(f".//{A}xfrm")
                if ln is None or xf is None:
                    continue
                ends = [(t, ln.find(f"{A}{t}")) for t in ("headEnd", "tailEnd")]
                ends = [(t, e) for t, e in ends
                        if e is not None and (e.get("type") or "none") != "none"]
                if not ends:
                    continue
                if drawings is None:
                    drawings = [d for d in page.get_drawings()
                                if d["type"] in ("f", "fs")]
                off, ext = xf.find(f"{A}off"), xf.find(f"{A}ext")
                ox, oy = int(off.get("x")) / 12700, int(off.get("y")) / 12700
                ecx, ecy = int(ext.get("cx")) / 12700, int(ext.get("cy")) / 12700
                lw = int(ln.get("w") or 9525) / 12700
                p0, p1 = [ox, oy], [ox + ecx, oy + ecy]
                if xf.get("flipH") == "1":
                    p0[0], p1[0] = p1[0], p0[0]
                if xf.get("flipV") == "1":
                    p0[1], p1[1] = p1[1], p0[1]
                rot = math.radians(int(xf.get("rot") or 0) / 60000.0)
                cx, cy = ox + ecx / 2, oy + ecy / 2

                def place(p):
                    dx, dy = p[0] - cx, p[1] - cy
                    return ((cx + dx * math.cos(rot) - dy * math.sin(rot)) * k,
                            (cy + dx * math.sin(rot) + dy * math.cos(rot)) * k)

                a, b = place(p0), place(p1)
                L = math.hypot(b[0] - a[0], b[1] - a[1])
                if L < 4:
                    continue
                ux, uy = (b[0] - a[0]) / L, (b[1] - a[1]) / L
                for tname, e in ends:
                    tip = a if tname == "headEnd" else b
                    # the filled path whose points cluster at this endpoint
                    best = None
                    for d in drawings:
                        pts = path_points(d)
                        if not pts or len(pts) > 200:
                            continue
                        cxp = sum(p[0] for p in pts) / len(pts)
                        cyp = sum(p[1] for p in pts) / len(pts)
                        dist = math.hypot(cxp - tip[0], cyp - tip[1])
                        if dist > 8 * lw * k + 4:
                            continue
                        if best is None or dist < best[0]:
                            best = (dist, pts)
                    if best is None:
                        continue
                    pts = best[1]
                    sgn = 1 if tname == "headEnd" else -1
                    along = [((p[0] - tip[0]) * ux + (p[1] - tip[1]) * uy) * sgn
                             for p in pts]
                    across = [-(p[0] - tip[0]) * uy + (p[1] - tip[1]) * ux
                              for p in pts]
                    rows[(e.get("type"), e.get("w") or "med",
                          e.get("len") or "med", round(lw, 2))].append((
                        (max(across) - min(across)) / k,
                        (max(along) - min(along)) / k,
                        min(along) / k,      # how far it reaches PAST the end
                        deck[:10], sn))

    print(f"{'type':9s} {'w':4s} {'len':4s} {'lw':>5s} {'n':>4s} | "
          f"{'across':>7s} {'along':>7s} {'past':>7s} | {'acr/lw':>6s} "
          f"{'aln/lw':>6s} {'pst/lw':>6s}  decks")
    for key in sorted(rows):
        v = rows[key]
        med = lambda i: sorted(x[i] for x in v)[len(v) // 2]
        acr, aln, pst = med(0), med(1), med(2)
        lw = key[3]
        decks = sorted({x[3] for x in v})
        print(f"{key[0]:9s} {key[1]:4s} {key[2]:4s} {lw:5.2f} {len(v):4d} | "
              f"{acr:7.3f} {aln:7.3f} {pst:7.3f} | {acr / lw:6.3f} "
              f"{aln / lw:6.3f} {pst / lw:6.3f}  {len(decks)}: {','.join(decks[:4])}")


if __name__ == "__main__":
    main()
