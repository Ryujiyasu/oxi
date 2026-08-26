# -*- coding: utf-8 -*-
"""Read the PATH (radial) gradient probe back and compare Oxi against PowerPoint.

Renders nothing. Point it at a probe that has already been exported by
`export_pptx_gradpath.py` and rendered by the pptx renderer into
`pipeline_data/pptx_probes/gradpath/oxi_png/`.

Usage: python tools/metrics/read_pptx_gradpath.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_probes" / "gradpath"
EMU_PT = 914400 / 72
DPI = 100
K = DPI / 72


def shape_of(img: np.ndarray, m: dict) -> np.ndarray:
    x0 = int(m["x_emu"] / EMU_PT * K)
    y0 = int(m["y_emu"] / EMU_PT * K)
    w = int(m["w_emu"] / EMU_PT * K)
    h = int(m["h_emu"] / EMU_PT * K)
    return img[y0 + 2 : y0 + h - 2, x0 + 2 : x0 + w - 2]


def radial(sub: np.ndarray, focus: tuple[float, float], bins: int = 16):
    """Mean level per bin of distance/|focus->farthest corner|."""
    h, w = sub.shape
    cx, cy = focus[0] * (w - 1), focus[1] * (h - 1)
    ys, xs = np.mgrid[0:h, 0:w]
    r = np.sqrt((xs - cx) ** 2 + (ys - cy) ** 2)
    r_max = max(
        np.hypot(cx - px, cy - py)
        for px in (0, w - 1)
        for py in (0, h - 1)
    )
    r = r / r_max
    out = []
    for i in range(bins):
        lo, hi = i / bins, (i + 1) / bins
        m = (r >= lo) & (r < hi)
        out.append(sub[m].mean() if m.sum() > 100 else np.nan)
    return out


def main() -> None:
    man = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    pdf = pymupdf.open(ROOT / "probe_gradpath.pdf")
    png_dir = ROOT / "oxi_png"
    print(f"{'arm':<20}{'mean|d|':>9}{'max|d|':>8}   PowerPoint focus / Oxi focus")
    for m in man:
        pix = pdf[m["slide"] - 1].get_pixmap(dpi=DPI)
        ppt = np.asarray(
            Image.frombytes("RGB", (pix.width, pix.height), pix.samples), float
        ).mean(axis=2) / 255.0
        p = png_dir / f"slide_s{m['slide']}.png"
        if not p.exists():
            print(f"{m['name']:<20}  (no Oxi render)")
            continue
        oxi = np.asarray(Image.open(p).convert("RGB"), float).mean(axis=2) / 255.0
        P, O = shape_of(ppt, m), shape_of(oxi, m)
        d = np.abs(O - P)
        # darkest point = the pos-0 stop, i.e. the focus
        def focus(a):
            iy, ix = np.unravel_index(np.argmin(a), a.shape)
            return ix / a.shape[1], iy / a.shape[0]
        fp, fo = focus(P), focus(O)
        print(
            f"{m['name']:<20}{d.mean():>9.4f}{d.max():>8.3f}   "
            f"({fp[0]:.2f},{fp[1]:.2f}) / ({fo[0]:.2f},{fo[1]:.2f})"
        )


if __name__ == "__main__":
    main()
