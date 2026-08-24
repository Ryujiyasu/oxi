# -*- coding: utf-8 -*-
"""Read the shape-gradient orientation probe back out of PowerPoint's PDF.

For each arm the shape's box is known in EMU, so the reader rasterises the
slide, takes the interior of that box (inset so the shape edge and its
antialiasing are out of the sample) and fits

    luminance(x, y) = c + gx * x + gy * y

by least squares.  `atan2(-gy, gx)` in degrees is then the direction the ramp
BRIGHTENS in, measured the way `a:lin@ang` measures (0 = to the right, growing
clockwise on screen), so it can be compared with the declared `ang` directly.
`r2` says how planar the sample really is -- a ramp that is not linear, or a
box the reader missed, shows up there rather than as a plausible wrong angle.

Usage: python tools/metrics/read_pptx_gradrot.py [--oxi TAG]

With --oxi the same fit is run against the Oxi PNGs of the same probe rendered
under `pipeline_data/pptx_probes/gradrot/oxi_<TAG>`, so the two engines' angles
sit in one table.
"""
from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_probes" / "gradrot"
DPI = 150
EMU_PER_PT = 12700
INSET = 0.12  # fraction of the box dropped on every side before fitting


def fit(img: np.ndarray, box_px: tuple[float, float, float, float]):
    x0, y0, w, h = box_px
    ix, iy = w * INSET, h * INSET
    a = int(round(x0 + ix))
    b = int(round(y0 + iy))
    c = int(round(x0 + w - ix))
    d = int(round(y0 + h - iy))
    a, b = max(a, 0), max(b, 0)
    c, d = min(c, img.shape[1]), min(d, img.shape[0])
    if c - a < 8 or d - b < 8:
        return None
    patch = img[b:d, a:c].astype(np.float64)
    lum = patch.mean(axis=2) if patch.ndim == 3 else patch
    hh, ww = lum.shape
    ys, xs = np.mgrid[0:hh, 0:ww]
    A = np.column_stack([np.ones(hh * ww), xs.ravel(), ys.ravel()])
    coef, *_ = np.linalg.lstsq(A, lum.ravel(), rcond=None)
    pred = A @ coef
    resid = lum.ravel() - pred
    ss_tot = ((lum.ravel() - lum.mean()) ** 2).sum()
    r2 = 1.0 - (resid**2).sum() / ss_tot if ss_tot > 0 else 0.0
    gx, gy = coef[1], coef[2]
    mag = math.hypot(gx, gy)
    # Screen y grows downward; a:lin@ang grows clockwise from "to the right",
    # which is the same sense, so gy is used as-is.
    ang = math.degrees(math.atan2(gy, gx)) % 360.0
    return ang, mag, r2


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--oxi", help="also fit Oxi PNGs rendered under oxi_<TAG>")
    args = ap.parse_args()

    arms = json.loads((OUT / "arms.json").read_text(encoding="utf-8"))
    pdf = OUT / "probe_gradrot.pdf"
    if not pdf.exists():
        sys.exit(f"missing {pdf} -- run export_pptx_gradrot.py first")
    doc = pymupdf.open(pdf)
    scale = DPI / 72.0

    oxi_dir = OUT / f"oxi_{args.oxi}" if args.oxi else None

    print(f"{'arm':<20} {'ang':>5} {'rot':>4} {'fH':>2} {'fV':>2} {'rws':>4} "
          f"| {'PPT':>7} {'r2':>5} | {'OXI':>7} {'r2':>5} | {'d':>7}")
    rows = []
    for rec in arms:
        page = doc[rec["slide"] - 1]
        pix = page.get_pixmap(dpi=DPI)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3]
        x, y, w, h = rec["box"]
        box_px = (x / EMU_PER_PT * scale, y / EMU_PER_PT * scale,
                  w / EMU_PER_PT * scale, h / EMU_PER_PT * scale)
        got = fit(img, box_px)
        ppt_ang, _, ppt_r2 = got if got else (float("nan"),) * 3

        oxi_ang = oxi_r2 = float("nan")
        if oxi_dir:
            cand = sorted(oxi_dir.glob(f"*s{rec['slide']}.png"))
            cand = [p for p in cand if p.stem.endswith(f"s{rec['slide']}")]
            if cand:
                oimg = np.asarray(Image.open(cand[0]).convert("RGB"))
                sx = oimg.shape[1] / pix.width
                obox = tuple(v * sx for v in box_px)
                g2 = fit(oimg, obox)
                if g2:
                    oxi_ang, _, oxi_r2 = g2

        delta = float("nan")
        if not math.isnan(oxi_ang):
            delta = (oxi_ang - ppt_ang + 180.0) % 360.0 - 180.0
        rws = rec["rotWithShape"] if rec["rotWithShape"] is not None else "-"
        print(f"{rec['arm']:<20} {rec['ang']:>5} {rec['rot']:>4} {rec['flipH']:>2} "
              f"{rec['flipV']:>2} {rws:>4} | {ppt_ang:7.2f} {ppt_r2:5.3f} | "
              f"{oxi_ang:7.2f} {oxi_r2:5.3f} | {delta:7.2f}")
        rows.append({**rec, "ppt_ang": ppt_ang, "ppt_r2": ppt_r2,
                     "oxi_ang": oxi_ang, "oxi_r2": oxi_r2, "delta": delta})
    (OUT / "measured.json").write_text(json.dumps(rows, indent=1), encoding="utf-8")
    print(f"\nwrote {OUT / 'measured.json'}")


if __name__ == "__main__":
    main()
