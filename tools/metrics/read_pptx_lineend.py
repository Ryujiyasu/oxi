# -*- coding: utf-8 -*-
"""Read the line-end repro: how big is each arrowhead, in line widths?

For every connector in `gen_pptx_lineend.py`'s grid, measure:
  stem   - the stroke's thickness at the middle of the line
  across - the decoration's widest extent perpendicular to the line
  along  - how far the decoration's ink runs along the line
  past   - how far that ink reaches BEYOND the declared endpoint, which is what
           says whether the head is centred on the end or sits behind it
All in points, and each also divided by the line width, because the DrawingML
size tokens are documented as multiples of it.

With `--oxi <png dir>` the same measurements are taken from Oxi's render of the
same deck and printed beside PowerPoint's, so the implementation can be checked
against the law it was derived from.

Usage:
    python tools/metrics/read_pptx_lineend.py [--pdf probe.pdf] [--oxi DIR]
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import numpy as np

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PROBE = Path(r"pipeline_data\pptx_probe\probe_lineend.pptx")
DPI = 600  # the heads are a few points across; measure them generously
SLIDE_W_PT = 13.333 * 72  # the width gen_pptx_lineend.py gives the repro


def pages_from_pdf(path: Path) -> tuple[dict, float]:
    import pymupdf
    pdf = pymupdf.open(path)
    scale = DPI / 72.0
    out = {}
    for i in range(pdf.page_count):
        pix = pdf[i].get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
        a = np.frombuffer(pix.samples, dtype=np.uint8)
        out[i + 1] = a.reshape(pix.height, pix.width, 3).mean(axis=2) < 160
    return out, scale


def pages_from_pngs(dirpath: Path, slide_w_pt: float) -> tuple[dict, float]:
    from PIL import Image
    out, scale = {}, None
    for p in sorted(dirpath.glob("*_s*.png")):
        n = int(p.stem.split("_s")[-1])
        a = np.asarray(Image.open(p).convert("L"))
        out[n] = a < 160
        scale = a.shape[1] / slide_w_pt
    if scale is None:
        raise SystemExit(f"no slide PNGs in {dirpath}")
    return out, scale


def measure(pages: dict, scale: float, row: dict) -> dict | None:
    im = pages.get(row["slide"])
    if im is None:
        return None
    h, w = im.shape
    lw = row["line_pt"]
    y = row["y_pt"] * scale
    x0, x1 = row["x0_pt"] * scale, row["x1_pt"] * scale
    # The widest head is 5 floored-widths across, so the window has to clear
    # 2.5 of them -- a window that clips reports the WINDOW, not the head, and
    # it reports the same wrong number for both arms, which reads as agreement.
    pad = max(24.0, 4.0 * max(lw, 2.0) * scale)
    ys = slice(max(0, int(y - pad)), min(h, int(y + pad)))
    stem = im[ys, int((x0 + x1) / 2)].sum() / scale
    lo = max(0, int(x0 - pad))
    win = im[ys, lo:int(x0 + pad)]
    # EXTENT per column, not the ink COUNT: a stealth head is concave, so its
    # widest column holds two corner tips and little between them, and a count
    # would report that gap as a narrower head.
    idx = np.arange(win.shape[0])[:, None]
    hi = np.where(win.any(axis=0), (idx * win).max(axis=0), 0)
    lo_i = np.where(win.any(axis=0),
                    np.where(win, idx, win.shape[0]).min(axis=0), 0)
    heights = np.where(win.any(axis=0), hi - lo_i + 1, 0)
    cols = np.nonzero(heights)[0]
    if not len(cols):
        return {"stem": stem, "across": 0.0, "along": 0.0, "past": 0.0}
    return {
        "stem": stem,
        "across": heights.max() / scale,
        "along": (cols.max() - cols.min() + 1) / scale,
        "past": (int(x0) - (cols.min() + lo)) / scale,
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", default=str(PROBE.with_suffix(".pdf")))
    ap.add_argument("--oxi", help="directory of Oxi's slide PNGs for the same deck")
    args = ap.parse_args()

    manifest = json.loads(PROBE.with_suffix(".json").read_text(encoding="utf-8"))
    ppt, ppt_scale = pages_from_pdf(Path(args.pdf))
    oxi = oxi_scale = None
    if args.oxi:
        oxi, oxi_scale = pages_from_pngs(Path(args.oxi), SLIDE_W_PT)

    head = (f"{'type':9s} {'w':4s} {'len':4s} {'lw':>5s} | {'stem':>5s} "
            f"{'across':>6s} {'along':>6s} {'past':>6s} | {'acr/lw':>6s} "
            f"{'aln/lw':>6s} {'pst/lw':>6s}")
    if oxi:
        head += f" || {'across':>6s} {'along':>6s} {'past':>6s}"
    print(head)
    for row in manifest:
        m = measure(ppt, ppt_scale, row)
        if m is None:
            continue
        lw = row["line_pt"]
        line = (f"{row['type']:9s} {row['w_tok']:4s} {row['len_tok']:4s} {lw:5.2f} | "
                f"{m['stem']:5.2f} {m['across']:6.2f} {m['along']:6.2f} {m['past']:6.2f} | "
                f"{m['across'] / lw:6.2f} {m['along'] / lw:6.2f} {m['past'] / lw:6.2f}")
        if oxi:
            o = measure(oxi, oxi_scale, row)
            line += (f" || {o['across']:6.2f} {o['along']:6.2f} {o['past']:6.2f}"
                     if o else "  || (missing)")
        print(line)


if __name__ == "__main__":
    main()
