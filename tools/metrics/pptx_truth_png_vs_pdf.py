# -*- coding: utf-8 -*-
"""Ask whether a deck's truth PDF raster is what PowerPoint actually shows.

Compares PowerPoint's own `Slide.Export` PNG (`pptx_slide_png.py`) against
pymupdf's raster of the same page of `ppt_pdf/<doc>.pdf`, and -- when the Oxi
cache holds it -- scores Oxi against BOTH, so a deck whose PDF is misleading can
be re-scored against the PNG instead.

    python tools/metrics/pptx_truth_png_vs_pdf.py 31 --slides 2 --panel

Prints per slide: SSIM(png, pdf), SSIM(oxi, pdf), SSIM(oxi, png).
`--panel` writes a 3-up PNG (PDF raster | PowerPoint PNG | absolute difference).
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image
from skimage.metrics import structural_similarity as ssim

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
SSIM_DIR = REPO / "pipeline_data" / "pptx_benchmark" / "ssim_pptx"
DPI = 150


def arr(img: Image.Image, size) -> np.ndarray:
    if img.size != size:
        img = img.resize(size, Image.LANCZOS)
    return np.asarray(img.convert("RGB")).astype(float)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("doc")
    ap.add_argument("--slides", default="")
    ap.add_argument("--panel", action="store_true")
    args = ap.parse_args()

    key = f"{int(args.doc):02d}"
    pdf = pymupdf.open(SSIM_DIR / "ppt_pdf" / f"{key}.pdf")
    png_dir = SSIM_DIR / "ppt_png" / key
    oxi_dir = SSIM_DIR / "oxi_png" / key
    want = [int(s) for s in args.slides.split(",") if s.strip()] or [
        int(p.stem.split("_s")[1]) for p in sorted(png_dir.glob("slide_s*.png"))
    ]

    for i in sorted(want):
        p = png_dir / f"slide_s{i}.png"
        if not p.exists():
            print(f"s{i}: no PowerPoint PNG (run pptx_slide_png.py)")
            continue
        pix = pdf[i - 1].get_pixmap(dpi=DPI)
        ref = np.asarray(Image.frombytes("RGB", (pix.width, pix.height), pix.samples)).astype(float)
        size = (pix.width, pix.height)
        ppt = arr(Image.open(p), size)
        row = [f"s{i}", f"png/pdf {ssim(ref, ppt, channel_axis=2, data_range=255):.4f}"]
        o = oxi_dir / f"slide_s{i}.png"
        if o.exists():
            oxi = arr(Image.open(o), size)
            row.append(f"oxi/pdf {ssim(ref, oxi, channel_axis=2, data_range=255):.4f}")
            row.append(f"oxi/png {ssim(ppt, oxi, channel_axis=2, data_range=255):.4f}")
        print("  ".join(row), flush=True)

        if args.panel:
            diff = np.abs(ref - ppt).mean(axis=2)
            heat = np.stack([255 - diff, 255 - diff, np.full_like(diff, 255.0)], axis=2)
            panel = np.concatenate([ref, ppt, heat], axis=1)
            out = SSIM_DIR / f"_truthcheck_{key}_s{i}.png"
            Image.fromarray(panel.clip(0, 255).astype(np.uint8)).resize(
                (panel.shape[1] // 3, panel.shape[0] // 3), Image.LANCZOS).save(out)
            print(f"    panel -> {out}", flush=True)


if __name__ == "__main__":
    main()
