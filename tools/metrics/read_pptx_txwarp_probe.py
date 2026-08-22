# -*- coding: utf-8 -*-
"""Read the txwarp probe: what size did PowerPoint fit each shape's text at?

Rasterises PowerPoint's own PDF, measures each shape's ink inside its declared
box, and compares against the face's glyph metrics so the fitting rule can be
named (fit the ink WIDTH, the ink HEIGHT, or the line box).

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
FONT_FILES = {
    "Arial": r"C:\Windows\Fonts\arial.ttf",
    "Segoe UI": r"C:\Windows\Fonts\segoeui.ttf",
    "Comic Sans MS": r"C:\Windows\Fonts\comic.ttf",
}


def glyph_box(face: str, text: str):
    """(ink width, ink height, yMin, yMax) of `text` in em, from the outlines."""
    from fontTools.ttLib import TTFont
    f = TTFont(FONT_FILES[face], lazy=True)
    upm = f["head"].unitsPerEm
    cmap = f.getBestCmap()
    glyf = f["glyf"]
    hmtx = f["hmtx"]
    x = 0.0
    x0 = y0 = None
    x1 = y1 = None
    for ch in text:
        g = cmap[ord(ch)]
        gl = glyf[g]
        adv = hmtx[g][0]
        if gl.numberOfContours != 0:
            gx0, gx1 = x + gl.xMin, x + gl.xMax
            x0 = gx0 if x0 is None else min(x0, gx0)
            x1 = gx1 if x1 is None else max(x1, gx1)
            y0 = gl.yMin if y0 is None else min(y0, gl.yMin)
            y1 = gl.yMax if y1 is None else max(y1, gl.yMax)
        x += adv
    return ((x1 - x0) / upm, (y1 - y0) / upm, y0 / upm, y1 / upm)


def main() -> None:
    manifest = json.loads((DIR / "txwarp_manifest.json").read_text(encoding="utf-8"))
    pdf = pymupdf.open(DIR / "deck.pdf")
    pages = {}
    for i in range(pdf.page_count):
        pix = pdf[i].get_pixmap(matrix=pymupdf.Matrix(300 / 72, 300 / 72), alpha=False)
        rgb = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        pages[i + 1] = np.asarray(rgb.convert("L"))
    print(f"{'arm':14s} {'face':14s} {'box':>12s} {'ink w x h':>15s} "
          f"{'S(w-fit)':>9s} {'S(h-fit)':>9s} {'base-top/S':>11s}")
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
        base = (ys.max() + y0) / sy
        gw, gh, ymin, ymax = glyph_box(row["face"], row["text"])
        s_w = iw / gw
        s_h = ih / gh
        print(f"{row['label']:14s} {row['face']:14s} "
              f"{row['w']:5.0f}x{row['h']:<5.0f} {iw:7.1f} x {ih:5.1f} "
              f"{s_w:9.1f} {s_h:9.1f} {(base - row['y']) / s_w:11.4f}")


if __name__ == "__main__":
    main()
