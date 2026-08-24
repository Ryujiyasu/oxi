# -*- coding: utf-8 -*-
"""Read the turned-text probe back out of PowerPoint's PDF.

PyMuPDF hands every span its writing direction as a unit vector, so the baseline
ANGLE is read rather than inferred: `atan2(-dir_y, dir_x)` in degrees, with PDF
y already flipped into screen sense by `get_text`. The span's origin is the
first glyph's baseline point, which is what the renderer has to reproduce.

Reported per arm:

    ppt_deg   the baseline direction PowerPoint drew
    ox, oy    the first glyph's baseline origin, in POINTS from the page corner
    lines     how many lines the paragraph broke into (block W's question)
    cx, cy    the centre of the red frame rectangle's ink, which is where the
              box actually sits -- the turning-centre question

Usage: python tools/metrics/read_pptx_textrot.py [--dump SLIDE]
"""
from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_probes" / "textrot"
EMU_PT = 12700


def frame_centre(page):
    """Centre of the red hairline rectangle, from the page's vector drawings."""
    xs, ys = [], []
    for d in page.get_drawings():
        col = d.get("color")
        if col is None:
            continue
        r, g, b = col
        if r > 0.8 and g < 0.2 and b < 0.2:
            rect = d["rect"]
            xs += [rect.x0, rect.x1]
            ys += [rect.y0, rect.y1]
    if not xs:
        return None
    return ((min(xs) + max(xs)) / 2.0, (min(ys) + max(ys)) / 2.0)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--dump", type=int, help="print the raw span dict for one slide")
    args = ap.parse_args()

    arms = json.loads((OUT / "arms.json").read_text(encoding="utf-8"))
    pdf = OUT / "probe_textrot.pdf"
    if not pdf.exists():
        sys.exit(f"missing {pdf} -- run export_pptx_textrot.py first")
    doc = pymupdf.open(pdf)

    x, y, w, h = arms[0]["box"]
    bx, by = x / EMU_PT, y / EMU_PT
    bw, bh = w / EMU_PT, h / EMU_PT
    print(f"box {bw:.1f} x {bh:.1f}pt at ({bx:.1f}, {by:.1f}); "
          f"unturned centre ({bx + bw / 2:.2f}, {by + bh / 2:.2f})")
    print(f"{'arm':<14} {'rot':>5} {'anc':>4} {'aln':>4} | {'ppt_deg':>8} "
          f"{'ox':>8} {'oy':>8} {'lines':>5} | {'frame_cx':>8} {'frame_cy':>8}")

    rows = []
    for rec in arms:
        page = doc[rec["slide"] - 1]
        d = page.get_text("dict")
        if args.dump == rec["slide"]:
            print(json.dumps(d, indent=1, default=str)[:4000])
        spans = [(ln, sp) for blk in d["blocks"] if blk.get("type") == 0
                 for ln in blk["lines"] for sp in ln["spans"] if sp["text"].strip()]
        if not spans:
            print(f"{rec['arm']:<14} {rec['rot']:>5} -- no text --")
            continue
        ln, sp = spans[0]
        dx, dy = ln.get("dir", (1.0, 0.0))
        deg = math.degrees(math.atan2(dy, dx)) % 360.0
        ox, oy = sp["origin"]
        nlines = len({round(s["origin"][0] * dy + s["origin"][1] * -dx, 1)
                      for _, s in spans})
        fc = frame_centre(page)
        fcx, fcy = fc if fc else (float("nan"), float("nan"))
        print(f"{rec['arm']:<14} {rec['rot']:>5} {rec['anchor']:>4} {rec['align']:>4} | "
              f"{deg:8.2f} {ox:8.2f} {oy:8.2f} {nlines:5d} | {fcx:8.2f} {fcy:8.2f}")
        rows.append({**rec, "ppt_deg": deg, "ox": ox, "oy": oy,
                     "lines": nlines, "frame_cx": fcx, "frame_cy": fcy})
    (OUT / "measured.json").write_text(json.dumps(rows, indent=1), encoding="utf-8")
    print(f"\nwrote {OUT / 'measured.json'}")
    ink_check(doc, arms, (bx + bw / 2, by + bh / 2))


def ink_check(doc, arms, centre):
    """Off-axis arms carry NO text objects -- PowerPoint exports text that is
    not axis-aligned as vector OUTLINES, so `get_text` returns nothing at all
    for rot 30 / 45 / 135 / -45. The ink is then the only instrument, and it is
    a sharp one: take the rot=0 arm's black pixels, turn them about the box
    centre by the arm's own angle, and the predicted bounding box has to be the
    one PowerPoint drew. There are no free parameters, so a wrong turning
    centre or a wrong sense misses by tens of points, not by a rounding."""
    import math

    import numpy as np

    DPI = 150
    sc = DPI / 72.0

    def black_ink(page):
        pix = page.get_pixmap(dpi=DPI)
        a = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3]
        m = a.max(axis=2) < 110          # text ink only; the frame is pure red
        ys, xs = np.nonzero(m)
        if xs.size == 0:
            return None, None
        return ((xs.min() / sc, ys.min() / sc, xs.max() / sc, ys.max() / sc),
                (xs / sc, ys / sc))

    base = next(r for r in arms if r["arm"] == "R_rot0")
    _, pts = black_ink(doc[base["slide"] - 1])
    if pts is None:
        print("no ink on the rot=0 arm; cannot run the ink check")
        return
    cx, cy = centre
    px, py = pts[0] - cx, pts[1] - cy
    print()
    print(f"INK CHECK -- rot=0 ink turned about ({cx:.1f}, {cy:.1f}) vs what PowerPoint drew")
    print(f"{'arm':<14} {'rot':>5} | {'predicted x0 y0 x1 y1':^33} | {'measured':^33} | {'max|d|':>7}")
    for rec in arms:
        if not rec["arm"].startswith("R_rot") or rec["rot"] == 0:
            continue
        th = math.radians(rec["rot"])
        c, s = math.cos(th), math.sin(th)
        qx = cx + px * c - py * s
        qy = cy + px * s + py * c
        pred = (qx.min(), qy.min(), qx.max(), qy.max())
        meas, _ = black_ink(doc[rec["slide"] - 1])
        if meas is None:
            print(f"{rec['arm']:<14} {rec['rot']:>5} | -- no ink --")
            continue
        d = max(abs(a - b) for a, b in zip(pred, meas))
        ps = " ".join(f"{v:7.2f}" for v in pred)
        ms = " ".join(f"{v:7.2f}" for v in meas)
        print(f"{rec['arm']:<14} {rec['rot']:>5} | {ps} | {ms} | {d:7.2f}")


if __name__ == "__main__":
    main()
