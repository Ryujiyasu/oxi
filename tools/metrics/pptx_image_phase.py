# -*- coding: utf-8 -*-
"""Who owns a raster image's sub-pixel offset: PowerPoint, the rasteriser, or Oxi.

S-IMGPHASE (2026-09-05, reverted) left the half-pixel image class open: on d39
s4 / d44 s13 the photo's interior sits ~0.5 output px away from the reference
while its edge agrees. The reference, though, is not PowerPoint -- it is
MuPDF's raster of PowerPoint's PDF, and MuPDF enlarges an image by its own
convention. Before touching the resampler again the offset has to be charged
to a side.

For every image the truth page draws, this tool:

  1. takes the image's exact placement from MuPDF's device
     (`get_image_info(xrefs=True)["transform"]` -- the effective ctm, no
     content-stream walking) and its own pixels (`Pixmap(doc, xref)`);
  2. builds an IDEAL raster at the scoring DPI: each output pixel's area is
     supersampled 4x4, every sample inverse-mapped through the ctm and read
     bilinearly between texel CENTRES (the continuous image, no convention);
  3. rasterises the same page with a SECOND rasteriser (pdfium) and reads the
     Oxi PNG of the slide;
  4. measures the sub-pixel shift of each raster against the ideal inside the
     image's interior (bbox shrunk by a margin), by upsampled phase correlation.

Reading: a side whose shift is ~0 agrees with the continuous image. If MuPDF
carries the offset and pdfium does not, the reference is the biased side and
Oxi must not chase it. Shifts are (dy, dx) in output pixels, positive = the
raster's content sits lower / further right than the ideal.

    python tools/metrics/pptx_image_phase.py d39:4 d44:13 [--tag s0929] [--dpi 150]
"""
from __future__ import annotations

import argparse
import glob
import sys
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image
from scipy.ndimage import map_coordinates
from skimage.registration import phase_cross_correlation

REPO = Path(__file__).resolve().parents[2]
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def gray(a: np.ndarray) -> np.ndarray:
    a = a.astype(np.float64)
    if a.ndim == 2:
        return a
    return a[..., :3] @ np.array([0.299, 0.587, 0.114])


def image_pixels(doc: pymupdf.Document, xref: int) -> np.ndarray:
    pix = pymupdf.Pixmap(doc, xref)
    if pix.colorspace is None or pix.colorspace.n not in (1, 3):
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    if pix.alpha:
        pix = pymupdf.Pixmap(pix, 0)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    return gray(arr)


def ideal_raster(img: np.ndarray, m: list[float], scale: float, box: tuple[int, int, int, int],
                 ss: int = 4) -> np.ndarray:
    """Continuous image through ctm `m` (unit square -> page pt), sampled over `box` (px)."""
    x0, y0, x1, y1 = box
    h, w = img.shape
    a, b, c, d, e, f = m
    det = a * d - b * c
    offs = (np.arange(ss) + 0.5) / ss
    acc = np.zeros((y1 - y0, x1 - x0))
    ys = np.arange(y0, y1)[:, None]
    xs = np.arange(x0, x1)[None, :]
    for oy in offs:
        for ox in offs:
            px = (xs + ox) / scale - e
            py = (ys + oy) / scale - f
            u = (d * px - c * py) / det
            v = (-b * px + a * py) / det
            u, v = np.broadcast_arrays(u, v)
            # texel centres sit at (i + 0.5) / n
            acc += map_coordinates(img, [v * h - 0.5, u * w - 0.5], order=1, mode="nearest")
    return acc / (ss * ss)


def pdfium_page(pdf_path: Path, index: int, scale: float) -> np.ndarray | None:
    try:
        import pypdfium2 as pdfium
    except ImportError:
        return None
    doc = pdfium.PdfDocument(str(pdf_path))
    bmp = doc[index].render(scale=scale)
    return gray(np.asarray(bmp.to_pil().convert("RGB")))


def shift(ref: np.ndarray, mov: np.ndarray) -> tuple[float, float, float]:
    s, err, _ = phase_cross_correlation(ref, mov, upsample_factor=40, normalization=None)
    # phase_cross_correlation returns the shift that registers mov onto ref;
    # the content of `mov` therefore sits at -s relative to ref.
    return -float(s[0]), -float(s[1]), float(err)


def find(pattern: str) -> Path:
    hits = glob.glob(pattern)
    if not hits:
        raise SystemExit(f"no match: {pattern}")
    return Path(hits[0])


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("slides", nargs="+", help="deck:slide, e.g. d39:4")
    ap.add_argument("--tag", default="s0929")
    ap.add_argument("--dpi", type=float, default=150.0)
    ap.add_argument("--margin", type=int, default=6, help="px trimmed off each side of the bbox")
    ap.add_argument("--min", type=int, default=48, help="skip interiors smaller than this (px)")
    args = ap.parse_args()
    scale = args.dpi / 72.0

    print(f"{'slide':10} {'xref':>5} {'box px':>22}  {'mupdf dy,dx':>14} {'pdfium dy,dx':>14} {'oxi dy,dx':>14} {'pdfium-mu':>14} {'oxi-mu':>14}")
    for spec in args.slides:
        deck, sn = spec.split(":")
        sn = int(sn)
        pdf_path = find(str(DEV / "pdf" / f"{deck}__*.pdf"))
        doc = pymupdf.open(pdf_path)
        page = doc[sn - 1]
        mu = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
        mu = gray(np.frombuffer(mu.samples, dtype=np.uint8).reshape(mu.height, mu.width, mu.n))
        pf = pdfium_page(pdf_path, sn - 1, scale)
        oxi_png = find(str(DEV / "oxi_png" / args.tag / f"{deck}__*" / f"slide_s{sn}.png"))
        ox = Image.open(oxi_png).convert("RGB")
        if ox.size != (mu.shape[1], mu.shape[0]):
            print(f"  {spec}: oxi {ox.size} != ref {mu.shape[::-1]} -- render at {args.dpi} dpi")
            continue
        ox = gray(np.asarray(ox))
        for info in page.get_image_info(xrefs=True):
            xref = info["xref"]
            if not xref:
                continue
            bx0, by0, bx1, by1 = (v * scale for v in info["bbox"])
            H, W = mu.shape
            box = (max(0, int(np.ceil(bx0)) + args.margin), max(0, int(np.ceil(by0)) + args.margin),
                   min(W, int(bx1) - args.margin), min(H, int(by1) - args.margin))
            if box[2] - box[0] < args.min or box[3] - box[1] < args.min:
                continue
            try:
                img = image_pixels(doc, xref)
            except Exception as exc:  # noqa: BLE001
                print(f"  {spec} xref {xref}: {exc}")
                continue
            ideal = ideal_raster(img, info["transform"], scale, box)
            x0, y0, x1, y1 = box
            # The ideal is only trustworthy for a plain placed image (no soft
            # mask, clip or tiling); when all three rasters agree with each
            # other but not with it, the ideal is the broken side. So the
            # rasters are ALSO reported against MuPDF directly.
            cells = []
            ref = mu[y0:y1, x0:x1]
            for name, arr in (("mupdf", mu), ("pdfium", pf), ("oxi", ox)):
                if arr is None or arr.shape != mu.shape:
                    cells.append(f"{'-':>14}")
                    continue
                dy, dx, _ = shift(ideal, arr[y0:y1, x0:x1])
                cells.append(f"{dy:+6.2f},{dx:+6.2f}")
            for name, arr in (("pdfium", pf), ("oxi", ox)):
                if arr is None or arr.shape != mu.shape:
                    cells.append(f"{'-':>14}")
                    continue
                dy, dx, _ = shift(ref, arr[y0:y1, x0:x1])
                cells.append(f"{dy:+6.2f},{dx:+6.2f}")
            print(f"{spec:10} {xref:5d} {str(box):>22}  " + " ".join(f"{c:>14}" for c in cells))


if __name__ == "__main__":
    main()
