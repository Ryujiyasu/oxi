# -*- coding: utf-8 -*-
"""Read the shadow repro: where does PowerPoint put a drop shadow, and how soft?

For each square in `gen_pptx_shadow.py`'s sweeps, measured off PowerPoint's own
PDF rasterised at 300 DPI:

  offset      the shadow ink's centre of mass minus the square's centre, which
              should come out as dist x (cos dir, -sin dir)
  reach       how far the penumbra extends past the square's edge along that
              direction (the 2% ink contour) -- the blurRad reading
  edge 10-90  the distance over which the penumbra goes from 10% to 90% of its
              own plateau, the part a box blur has to match
  plateau     the darkest the shadow gets, which is where `alpha` shows

The squares are black on white and spaced three side-lengths apart, so a
scanline reads one shadow at a time and "ink" is unambiguous.

Usage: python tools/metrics/read_pptx_shadow.py [probe.pdf]
"""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

import numpy as np
import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PROBE = Path(r"pipeline_data\pptx_probes\shadow\probe_shadow.pptx")
DPI = 300
SLIDE_W_PT = 13.333 * 72


def main() -> None:
    pdf_path = Path(sys.argv[1]) if len(sys.argv) > 1 else PROBE.with_suffix(".pdf")
    manifest = json.loads(PROBE.with_suffix(".json").read_text(encoding="utf-8"))
    pdf = pymupdf.open(pdf_path)
    pages: dict[int, np.ndarray] = {}
    scale = None
    for row in manifest:
        s = row["slide"]
        if s not in pages:
            page = pdf[s - 1]
            k = DPI / 72.0
            pix = page.get_pixmap(matrix=pymupdf.Matrix(k, k), alpha=False)
            a = np.frombuffer(pix.samples, dtype=np.uint8)
            pages[s] = a.reshape(pix.height, pix.width, 3).mean(axis=2)
            scale = pix.width / SLIDE_W_PT
        im = pages[s]
        px = lambda v: v * scale                       # noqa: E731
        x0, y0 = px(row["x_pt"]), px(row["y_pt"])
        side = px(row["side_pt"])
        pad = px(60.0)
        xa, xb = int(x0 - pad), int(x0 + side + pad)
        ya, yb = int(y0 - pad), int(y0 + side + pad)
        win = im[max(0, ya):yb, max(0, xa):xb]
        # "ink" = anything darker than white; the square itself is masked out
        ink = (255.0 - win) / 255.0
        yy, xx = np.mgrid[0:win.shape[0], 0:win.shape[1]]
        sq_x0, sq_y0 = x0 - max(0, xa), y0 - max(0, ya)
        inside = ((xx >= sq_x0 - 1) & (xx <= sq_x0 + side + 1)
                  & (yy >= sq_y0 - 1) & (yy <= sq_y0 + side + 1))
        shade = np.where(inside, 0.0, ink)
        total = shade.sum()
        if total < 1e-6:
            print(f'{row["sweep"]:6s} blur={row["blur_pt"]:5.2f} '
                  f'dist={row["dist_pt"]:5.2f} dir={row["dir_deg"]:5.1f} '
                  f'alpha={row["alpha_pc"]:5.1f}  NO SHADOW INK')
            continue
        cx = (shade * xx).sum() / total
        cy = (shade * yy).sum() / total
        dx = (cx - (sq_x0 + side / 2)) / scale
        dy = (cy - (sq_y0 + side / 2)) / scale
        # profile along the declared direction, from the square's centre out
        ang = math.radians(row["dir_deg"])
        ux, uy = math.cos(ang), math.sin(ang)
        prof = []
        for t in np.arange(0.0, pad + side, 0.5):
            x = sq_x0 + side / 2 + ux * t
            y = sq_y0 + side / 2 + uy * t
            xi, yi = int(round(x)), int(round(y))
            if 0 <= xi < win.shape[1] and 0 <= yi < win.shape[0]:
                prof.append((t, ink[yi, xi]))
        outside = [(t, v) for t, v in prof
                   if t > side / 2 + max(1.0, px(1.0))]
        peak = max((v for _, v in outside), default=0.0)
        reach = max((t for t, v in outside if v > 0.02), default=0.0)
        lo = next((t for t, v in outside if v <= 0.9 * peak), 0.0)
        hi = next((t for t, v in outside if v <= 0.1 * peak), 0.0)
        print(f'{row["sweep"]:6s} blur={row["blur_pt"]:5.2f} '
              f'dist={row["dist_pt"]:5.2f} dir={row["dir_deg"]:5.1f} '
              f'alpha={row["alpha_pc"]:5.1f} | offset=({dx:6.2f},{dy:6.2f})pt '
              f'reach={(reach - side / 2) / scale:6.2f}pt '
              f'edge10-90={(hi - lo) / scale:5.2f}pt plateau={peak:5.3f}')


if __name__ == "__main__":
    main()
