# -*- coding: utf-8 -*-
"""Check the line-end sizes on the dev corpus's REAL connectors, both arms.

The repro (`gen_pptx_lineend.py`) proves the law; this proves the renderer
obeys it on documents nobody authored for the purpose. For every decorated
connector the slides declare, it measures the head's extent across the line by
walking outward from the line's centre until the stroke colour stops matching
-- a bounded run, so it cannot leak into a same-coloured neighbour the way a
flood fill does -- and groups the medians by (type, w, line width) beside the
size the law predicts.

Run it once with no argument for PowerPoint's own render and once with a render
tag for Oxi's, and the two tables should agree:

    python tools/metrics/read_pptx_lineend_corpus.py
    python tools/metrics/read_pptx_lineend_corpus.py head
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

from PIL import Image

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
DEV = Path(r"pipeline_data\pptx_benchmark\dev")
FACTOR = {"sm": 2.0, "med": 3.0, "lg": 5.0}
FLOOR = 2.0


def endpoints(xfrm):
    """The connector's two ends in slide EMU, after its flips and rotation."""
    off, ext = xfrm.find(f"{A}off"), xfrm.find(f"{A}ext")
    ox, oy = int(off.get("x")), int(off.get("y"))
    ecx, ecy = int(ext.get("cx")), int(ext.get("cy"))
    p0, p1 = [ox, oy], [ox + ecx, oy + ecy]
    if xfrm.get("flipH") == "1":
        p0[0], p1[0] = p1[0], p0[0]
    if xfrm.get("flipV") == "1":
        p0[1], p1[1] = p1[1], p0[1]
    rot = math.radians(int(xfrm.get("rot") or 0) / 60000.0)
    cx, cy = ox + ecx / 2, oy + ecy / 2

    def place(p):
        dx, dy = p[0] - cx, p[1] - cy
        return (cx + dx * math.cos(rot) - dy * math.sin(rot),
                cy + dx * math.sin(rot) + dy * math.cos(rot))

    return place(p0), place(p1)


def measure_slide(png: Path, root, sw: int, sh: int, rows: list, deck: str):
    img = None
    for sp in root.iter(f"{P}cxnSp"):
        ln, xfrm = sp.find(f".//{A}ln"), sp.find(f".//{A}xfrm")
        if ln is None or xfrm is None:
            continue
        ends = [(t, ln.find(f"{A}{t}")) for t in ("headEnd", "tailEnd")]
        ends = [(t, e) for t, e in ends
                if e is not None and (e.get("type") or "none") != "none"]
        if not ends:
            continue
        if img is None:
            img = Image.open(png).convert("RGB")
            iw, ih = img.size
            px = img.load()
            pt_per_px = (sw / 914400 * 72) / iw
        lw = int(ln.get("w") or 9525) / 12700
        (ax, ay), (bx, by) = endpoints(xfrm)
        a = (ax * iw / sw, ay * ih / sh)
        b = (bx * iw / sw, by * ih / sh)
        length = math.hypot(b[0] - a[0], b[1] - a[1])
        if length < 12:
            continue
        ux, uy = (b[0] - a[0]) / length, (b[1] - a[1]) / length
        vx, vy = -uy, ux

        def at(t, s):
            x, y = a[0] + ux * t + vx * s, a[1] + uy * t + vy * s
            xi, yi = int(round(x)), int(round(y))
            return px[xi, yi] if 0 <= xi < iw and 0 <= yi < ih else None

        mid = length / 2
        bg = at(mid, 25) or at(mid, -25)
        if bg is None:
            continue
        want = max(
            (c for c in (at(mid, k * 0.5) for k in range(-10, 11)) if c),
            key=lambda c: sum(abs(c[i] - bg[i]) for i in range(3)), default=None)
        if want is None or sum(abs(want[i] - bg[i]) for i in range(3)) < 40:
            continue

        def near(c, tol=30):
            return c is not None and all(abs(c[i] - want[i]) <= tol
                                         for i in range(3))

        def across(t, step=0.5, reach=40):
            """Width of the stroke-coloured run across the line, or None if the
            run never ends -- then it has wandered into a same-coloured
            neighbour and the number would be fiction."""
            if not near(at(t, 0.0)):
                return 0.0
            lo = hi = 0.0
            k = step
            while k < reach and near(at(t, -k)):
                lo = k
                k += step
            if lo >= reach - step:
                return None
            k = step
            while k < reach and near(at(t, k)):
                hi = k
                k += step
            if hi >= reach - step:
                return None
            return lo + hi + 1.0

        if not across(mid):
            continue
        # A pointed head's widest place is a whole head-length back from the
        # tip, so sweep inward rather than sampling only at the endpoint.
        reach = 12 * lw / pt_per_px + 6
        steps = [i * 0.5 for i in range(-6, int(reach * 2) + 1)]
        for tname, e in ends:
            seen = [across(t if tname == "headEnd" else length - t)
                    for t in steps]
            if any(s is None for s in seen) or not max(seen):
                continue
            rows.append(((e.get("type"), e.get("w") or "med", round(lw, 2)),
                         max(seen) * pt_per_px, deck[:10]))


def main() -> None:
    tag = sys.argv[1] if len(sys.argv) > 1 else None
    rows: list = []
    for deck_path in sorted(glob.glob(str(DEV / "pptx" / "*.pptx"))):
        deck = Path(deck_path).stem
        zf = zipfile.ZipFile(deck_path)
        try:
            pres = ET.fromstring(zf.read("ppt/presentation.xml"))
        except KeyError:
            continue
        sz = pres.find(f"{P}sldSz")
        sw, sh = int(sz.get("cx")), int(sz.get("cy"))
        for name in zf.namelist():
            m = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
            if not m:
                continue
            raw = zf.read(name)
            if b"End" not in raw:
                continue
            sn = int(m.group(1))
            png = (DEV / "oxi_png" / tag / deck / f"slide_s{sn}.png") if tag \
                else (DEV / "ppt_png" / deck / f"p{sn}.png")
            if png.exists():
                measure_slide(png, ET.fromstring(raw), sw, sh, rows, deck)

    by = defaultdict(list)
    for key, span, _ in rows:
        by[key].append(span)
    print(f"arm = {tag or 'PowerPoint'}")
    print(f"{'type':9s} {'w':4s} {'lw':>5s} {'n':>4s} {'median':>7s} "
          f"{'predicted':>9s}  ratio  decks")
    for key in sorted(by):
        v = sorted(by[key])
        med = v[len(v) // 2]
        pred = FACTOR[key[1]] * max(key[2], FLOOR)
        decks = {d for k, _, d in rows if k == key}
        print(f"{key[0]:9s} {key[1]:4s} {key[2]:5.2f} {len(v):4d} {med:7.2f} "
              f"{pred:9.2f}  {med / pred:5.3f}  {len(decks)}")


if __name__ == "__main__":
    main()
