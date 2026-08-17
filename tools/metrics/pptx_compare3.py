# -*- coding: utf-8 -*-
"""3-panel slide comparison: PowerPoint | Oxi | diff heat.

Usage:
    python tools/metrics/pptx_compare3.py d01 3 --tag cg [--tag2 head] [--out X.png]

With --tag2 the middle panel is the --tag render and the right panel becomes
the A/B difference between the two Oxi arms instead of the PowerPoint diff.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image

REPO_ROOT = Path(__file__).resolve().parents[2]
DEV = REPO_ROOT / "pipeline_data" / "pptx_benchmark" / "dev"
DPI = 150


def deck_dir(deck: str) -> str:
    hits = sorted(DEV.joinpath("pptx").glob(f"{deck}*.pptx"))
    if not hits:
        sys.exit(f"no deck matching {deck}")
    return hits[0].stem


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("deck")
    ap.add_argument("slide", type=int)
    ap.add_argument("--tag", default="head")
    ap.add_argument("--tag2", default=None)
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    name = deck_dir(args.deck)
    pdf = pymupdf.open(DEV / "pdf" / f"{name}.pdf")
    pix = pdf[args.slide - 1].get_pixmap(matrix=pymupdf.Matrix(DPI / 72, DPI / 72), alpha=False)
    ref = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3]

    def load(tag: str) -> np.ndarray:
        png = DEV / "oxi_png" / tag / name / f"slide_s{args.slide}.png"
        img = Image.open(png).convert("RGB")
        if img.size != (ref.shape[1], ref.shape[0]):
            img = img.resize((ref.shape[1], ref.shape[0]), Image.LANCZOS)
        return np.asarray(img)

    mid = load(args.tag)
    right_src = load(args.tag2) if args.tag2 else ref
    diff = np.abs(mid.astype(int) - right_src.astype(int)).sum(axis=2)
    heat = np.zeros_like(mid)
    heat[..., 0] = np.clip(diff, 0, 255)
    heat[..., 2] = 255 - np.clip(diff, 0, 255)

    panel = np.concatenate([ref, mid, heat], axis=1)
    out = Path(args.out) if args.out else DEV / f"_cmp_{args.deck}_s{args.slide}_{args.tag}.png"
    Image.fromarray(panel).save(out)
    print(f"wrote {out}  ({panel.shape[1]}x{panel.shape[0]})")


if __name__ == "__main__":
    main()
