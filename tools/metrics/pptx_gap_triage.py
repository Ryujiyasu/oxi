# -*- coding: utf-8 -*-
"""Is this slide's SSIM gap a DEFECT, or is it anti-aliasing and reference noise?

A low SSIM says a slide differs; it does not say the difference is Oxi's fault
or worth chasing. Three of the dev corpus's worst slides turned out not to be
defects at all, each for a different reason, and each cost a session to find
out. This asks the question directly, per slide:

  edge share      how much of the error sits within 2px of an edge the
                  REFERENCE itself has. Anti-aliasing lives here. A slide whose
                  error is mostly edge share has no misplaced element -- it is
                  drawing the same shapes with a different rasteriser.

  flat blobs      connected runs of flat-area pixels that are badly wrong
                  (>40). THIS is what a real defect looks like: a missing fill,
                  a shape in the wrong place, a colour computed wrongly. A
                  handful of tiny blobs is AA spill; a big one is a bug.

  reference grain whether the reference's flat regions carry variance the
                  source does not. PowerPoint's PDF export re-encodes bitmaps
                  with DCT, so a correct flat render is penalised
                  ([[pptx-reference-jpeg-grain]]).

  best shift      the integer (dy, dx) that minimises the error. A non-zero
                  shift on one slide of a deck is a layout bug; the same shift
                  on every slide is a canvas offset.

Worked examples (2026-08-26):
  d28 s12  edge 21%  grain YES  -> reference JPEG noise, Oxi matches the source
  d09 s3   edge 61%  blobs max 24px  shift (0,1) on 1 of 15 slides
           -> anti-aliasing on vector art, plus a right-aligned block displaced
              1px because Oxi's line is 0.08% wide
  d15 s2   -- a wrap difference, which this does not diagnose: it shows up as
           large flat blobs where a word moved, so blobs alone do not prove a
           painting bug. Read the blob bbox before concluding.

Usage:
    python tools/metrics/pptx_gap_triage.py d09:3 d24:1 d32:6
    python tools/metrics/pptx_gap_triage.py d09          # worst slide of a deck
"""
from __future__ import annotations

import glob
import sys
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image
from scipy import ndimage

REPO = Path(__file__).resolve().parents[2]
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def load(deck: str, slide: int, tag: str):
    pdfs = sorted(glob.glob(str(DEV / "pdf" / f"{deck}__*.pdf")))
    pngs = glob.glob(str(DEV / "oxi_png" / tag / f"{deck}__*"))
    if not pdfs or not pngs:
        return None
    doc = pymupdf.open(pdfs[0])
    if slide > len(doc):
        return None
    pix = doc[slide - 1].get_pixmap(dpi=150)
    ref = np.asarray(Image.frombytes("RGB", (pix.width, pix.height), pix.samples)).astype(float)
    p = Path(pngs[0]) / f"slide_s{slide}.png"
    if not p.exists():
        return None
    o = Image.open(p).convert("RGB")
    if o.size != (pix.width, pix.height):
        o = o.resize((pix.width, pix.height), Image.LANCZOS)
    return ref, np.asarray(o).astype(float)


def triage(deck: str, slide: int, tag: str) -> None:
    got = load(deck, slide, tag)
    if got is None:
        print(f"{deck} s{slide}: missing reference or render (tag {tag})")
        return
    ref, oxi = got
    err = np.abs(ref - oxi).mean(axis=2)
    g = ref.mean(axis=2)
    edge = (np.abs(np.gradient(g, axis=1)) + np.abs(np.gradient(g, axis=0))) > 12
    edge = ndimage.binary_dilation(edge, iterations=2)
    flat_bad = (err > 40) & (~edge)
    lab, n = ndimage.label(flat_bad)
    sizes = ndimage.sum(flat_bad, lab, range(1, n + 1)) if n else np.array([0])
    # reference grain: variance in regions Oxi paints perfectly flat
    from collections import Counter
    common = Counter(map(tuple, oxi.reshape(-1, 3).astype(int))).most_common(1)[0]
    mask = np.all(oxi.astype(int) == np.array(common[0]), axis=2)
    grain = ref[mask].std(axis=0).mean() if mask.sum() > 5000 else float("nan")
    best = None
    for dy in (-2, -1, 0, 1, 2):
        for dx in (-2, -1, 0, 1, 2):
            e = np.abs(ref - np.roll(np.roll(oxi, dy, 0), dx, 1))[6:-6, 6:-6].mean()
            if best is None or e < best[0]:
                best = (e, dy, dx)
    share = err[edge].sum() / err.sum() * 100 if err.sum() else 0
    print(f"{deck} s{slide}:  mean|err| {err.mean():.2f}")
    print(f"   edge share      {share:5.1f}%  (edges are {edge.mean() * 100:.1f}% of area)")
    print(f"   flat blobs      {n:5d}  largest {int(sizes.max()):d} px"
          f"{'   <- look here' if sizes.max() > 400 else '   (all small: AA spill)'}")
    print(f"   reference grain {grain:5.2f}  on {int(mask.sum())} px Oxi paints flat"
          f"{'   <- reference noise' if grain > 2 else ''}")
    print(f"   best shift      dy={best[1]} dx={best[2]}  err {best[0]:.2f} (at 0,0 {np.abs(ref - oxi)[6:-6, 6:-6].mean():.2f})")


def main() -> None:
    args = sys.argv[1:]
    tag = "s0826a"
    if not args:
        sys.exit("usage: pptx_gap_triage.py d09:3 [d24:1 ...]")
    for a in args:
        if ":" in a:
            deck, s = a.split(":")
            triage(deck, int(s), tag)
        else:
            triage(a, 1, tag)


if __name__ == "__main__":
    main()
