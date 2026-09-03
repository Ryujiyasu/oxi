# -*- coding: utf-8 -*-
"""`pptx_defect_rank.py`'s arithmetic, one process per deck.

The ranker scores every slide twice -- raw and blurred -- over the whole blind
set, which is about nine hundred 3000x1688 comparisons. None of it touches the
renderer, so the only reason it ran on one core is that it was written beside a
tool that has to (`pptx_render_not_parallel_safe` governs the RENDERER).

Same numbers, same meaning: `defect = 1 - SSIM(blur(ref), blur(oxi))`, and
`explained` is the share of the raw gap the blur closes -- high means grain or
antialiasing, low means something is actually in the wrong place.

    python tools/metrics/pptx_defect_rank_par.py [--limit N] [--jobs N]
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image, ImageFilter
from skimage.metrics import structural_similarity

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
SS = ROOT / "ssim_pptx"
OUT = SS / "_defect_rank_par.json"
BLUR = 1.0

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def ssim(a: np.ndarray, b: np.ndarray) -> float:
    return float(structural_similarity(a, b, channel_axis=2, data_range=255))


def score_deck(doc: str) -> list[dict]:
    pdf_path = SS / "ppt_pdf" / f"{doc}.pdf"
    png_dir = SS / "oxi_png" / doc
    if not pdf_path.exists() or not png_dir.is_dir():
        return []
    out = []
    pdf = pymupdf.open(pdf_path)
    for png in sorted(png_dir.glob("slide_s*.png"), key=lambda p: int(p.stem.split("_s")[1])):
        idx = int(png.stem.split("_s")[1])
        if idx > len(pdf):
            continue
        oxi = Image.open(png).convert("RGB")
        pg = pdf[idx - 1]
        k = oxi.width / pg.rect.width
        pix = pg.get_pixmap(matrix=pymupdf.Matrix(k, k))
        ref = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        if ref.size != oxi.size:
            ref = ref.resize(oxi.size, Image.LANCZOS)
        r, o = np.asarray(ref, float), np.asarray(oxi, float)
        raw = ssim(r, o)
        rb = np.asarray(ref.filter(ImageFilter.GaussianBlur(BLUR)), float)
        ob = np.asarray(oxi.filter(ImageFilter.GaussianBlur(BLUR)), float)
        blur = ssim(rb, ob)
        out.append({
            "doc": doc, "slide": idx,
            "ssim": round(raw, 5), "ssim_blur": round(blur, 5),
            "defect": round(1 - blur, 5),
            "explained": round((blur - raw) / max(1 - raw, 1e-6), 4),
            "mean_err": round(float(np.abs(r - o).mean()), 3),
            "heavy": round(float((np.abs(r - o).mean(axis=2) > 40).mean() * 100), 3),
        })
    pdf.close()
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=20)
    ap.add_argument("--jobs", type=int, default=0)
    args = ap.parse_args()
    man = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    docs = [f"{m['idx']:02d}" for m in man]
    jobs = args.jobs or max(1, (os.cpu_count() or 4) - 1)
    rows: list[dict] = []
    with ProcessPoolExecutor(max_workers=jobs) as pool:
        for got in pool.map(score_deck, docs):
            rows.extend(got)
            if got:
                print(f"{got[0]['doc']}: {len(got)} slides", flush=True)
    OUT.write_text(json.dumps(rows, indent=1), encoding="utf-8")
    rows.sort(key=lambda r: -r["defect"])
    print(f"\nworst {args.limit} slides by BLURRED residual (blur r={BLUR}px)")
    print(f"{'slide':<10}{'defect':>9}{'ssim':>9}{'blurred':>9}{'explained':>11}"
          f"{'mean|err|':>11}{'heavy%':>9}")
    for r in rows[:args.limit]:
        print(f"{r['doc']}/{r['slide']:<7}{r['defect']:>9.4f}{r['ssim']:>9.4f}"
              f"{r['ssim_blur']:>9.4f}{r['explained'] * 100:>10.1f}%"
              f"{r['mean_err']:>11.2f}{r['heavy']:>9.2f}")
    print(f"\nwrote {OUT}  ({len(rows)} slides)")


if __name__ == "__main__":
    main()
