# -*- coding: utf-8 -*-
"""Read the txwarp probe: how did PowerPoint fit each shape's text to its box?

Rasterises PowerPoint's own PDF and measures each shape's ink against the box
the shape declares: where the ink starts as a fraction of the box (dx/w, dy/h)
and how much of the box it fills (w ratio, h ratio). Those four numbers say
directly whether the text is fitted to the box's width, its height, or neither,
without needing the face's own metrics as an intermediary.

Usage: python tools/metrics/read_pptx_txwarp_probe.py
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

DIR = Path(r"pipeline_data\pptx_probes\txwarp").resolve()


def main() -> None:
    manifest = json.loads((DIR / "txwarp_manifest.json").read_text(encoding="utf-8"))
    pdf = pymupdf.open(DIR / "deck.pdf")
    pages = {}
    for i in range(pdf.page_count):
        pix = pdf[i].get_pixmap(matrix=pymupdf.Matrix(300 / 72, 300 / 72), alpha=False)
        rgb = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        pages[i + 1] = np.asarray(rgb.convert("L"))
    print(f"{'arm':14s} {'face':14s} {'box':>12s} {'ink w x h':>15s} "
          f"{'dx/w':>7s} {'dy/h':>7s} {'w ratio':>8s} {'h ratio':>8s}")
    for row in manifest:
        im = pages[row["slide"]]
        h, w = im.shape
        sx, sy = w / 720.0, h / 540.0
        pad = 40
        x0 = max(int((row["x"] - pad) * sx), 0)
        x1 = min(int((row["x"] + row["w"] + pad) * sx), w)
        y0 = max(int((row["y"] - pad) * sy), 0)
        y1 = min(int((row["y"] + row["h"] + pad) * sy), h)
        win = im[y0:y1, x0:x1]
        mask = win < 160
        ys, xs = np.nonzero(mask)
        if not len(ys):
            print(f"{row['label']:14s} {row['face']:14s} no ink")
            continue
        iw = (xs.max() - xs.min() + 1) / sx
        ih = (ys.max() - ys.min() + 1) / sy
        ink_x = (xs.min() + x0) / sx
        ink_y = (ys.min() + y0) / sy
        print(f"{row['label']:14s} {row['face']:14s} "
              f"{row['w']:5.0f}x{row['h']:<5.0f} {iw:7.1f} x {ih:5.1f} "
              f"{(ink_x - row['x']) / row['w']:7.3f} "
              f"{(ink_y - row['y']) / row['h']:7.3f} "
              f"{iw / row['w']:8.3f} {ih / row['h']:8.3f}")


if __name__ == "__main__":
    main()
