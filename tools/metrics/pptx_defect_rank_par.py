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


def picture_mask(pg, size) -> np.ndarray | None:
    """Where the truth page places a raster image, in the render's pixels.

    The corpus's worst slides by raw or blurred residual are, over and over,
    the ones carrying a photograph: the truth PDF re-encodes what the deck
    embedded, so a photo drawn perfectly still differs by a few units per pixel
    over a large area (`pptx_reference_jpeg_grain`). Blur does not remove it --
    it is not antialiasing, it is different grain -- so the blurred ranking is
    still led by pictures. Masking them out asks the question the ranker was
    built for: where is the ENGINE wrong.

    Returns a boolean array that is True where the page draws no image, or
    None when the page has none at all.
    """
    rects = []
    for info in pg.get_images(full=True):
        try:
            rects.extend(pg.get_image_rects(info[0]))
        except Exception:
            continue
    if not rects:
        return None
    w, h = size
    kx, ky = w / pg.rect.width, h / pg.rect.height
    keep = np.ones((h, w), dtype=bool)
    for r in rects:
        x0 = max(0, int(r.x0 * kx) - 1)
        y0 = max(0, int(r.y0 * ky) - 1)
        x1 = min(w, int(r.x1 * kx) + 1)
        y1 = min(h, int(r.y1 * ky) + 1)
        if x1 > x0 and y1 > y0:
            keep[y0:y1, x0:x1] = False
    return keep


def masked_mean(smap: np.ndarray, keep: np.ndarray | None) -> float:
    """The SSIM map's mean over the pixels `keep` marks.

    Averaging the map over a subset is not the same number as SSIM of a cropped
    image, and it is the one that composes: every slide keeps its own irregular
    set of non-picture pixels. The border skimage excludes from its own scalar
    (half a window) is excluded here too, so the two numbers stay comparable.
    """
    m = smap.mean(axis=2)
    pad = 3  # (win_size 7 - 1) // 2
    m = m[pad:-pad, pad:-pad]
    if keep is None:
        return float(m.mean())
    k = keep[pad:-pad, pad:-pad]
    if not k.any():
        return float("nan")
    return float(m[k].mean())


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
        # One pass, both numbers: the scalar skimage returns IS the mean of the
        # map it returns, so asking for the map costs nothing beyond memory.
        blur, smap = structural_similarity(
            rb, ob, channel_axis=2, data_range=255, full=True
        )
        keep = picture_mask(pg, oxi.size)
        nopic = masked_mean(smap, keep)
        area = 100.0 if keep is None else round(float(keep.mean()) * 100, 1)
        out.append({
            "doc": doc, "slide": idx,
            "ssim": round(raw, 5), "ssim_blur": round(blur, 5),
            "defect": round(1 - blur, 5),
            "defect_nopic": None if nopic != nopic else round(1 - nopic, 5),
            "open_area": area,
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
    ap.add_argument("--decks", default="", help="comma-separated indices, else all")
    args = ap.parse_args()
    man = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    docs = [f"{m['idx']:02d}" for m in man]
    if args.decks:
        want = {d.strip().zfill(2) for d in args.decks.split(",") if d.strip()}
        docs = [d for d in docs if d in want]
    jobs = args.jobs or max(1, (os.cpu_count() or 4) - 1)
    rows: list[dict] = []
    with ProcessPoolExecutor(max_workers=jobs) as pool:
        for got in pool.map(score_deck, docs):
            rows.extend(got)
            if got:
                print(f"{got[0]['doc']}: {len(got)} slides", flush=True)
    if args.decks and OUT.exists():
        # A partial run keeps what the others measured, so the ranking can be
        # rebuilt a few decks at a time without losing the rest.
        done = {r["doc"] for r in rows}
        try:
            old = json.loads(OUT.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            old = []
        rows = rows + [r for r in old if r["doc"] not in done]
    OUT.write_text(json.dumps(rows, indent=1), encoding="utf-8")
    rows.sort(key=lambda r: -r["defect"])
    print(f"\nworst {args.limit} slides by BLURRED residual (blur r={BLUR}px)")
    print(f"{'slide':<10}{'defect':>9}{'ssim':>9}{'blurred':>9}{'explained':>11}"
          f"{'mean|err|':>11}{'heavy%':>9}")
    for r in rows[:args.limit]:
        print(f"{r['doc']}/{r['slide']:<7}{r['defect']:>9.4f}{r['ssim']:>9.4f}"
              f"{r['ssim_blur']:>9.4f}{r['explained'] * 100:>10.1f}%"
              f"{r['mean_err']:>11.2f}{r['heavy']:>9.2f}")
    keeps = [r for r in rows if r.get("defect_nopic") is not None]
    keeps.sort(key=lambda r: -r["defect_nopic"])
    print(f"\nworst {args.limit} slides OUTSIDE the pictures (the engine's own residual)")
    print(f"{'slide':<10}{'nopic':>9}{'defect':>9}{'open%':>8}{'mean|err|':>11}{'heavy%':>9}")
    for r in keeps[:args.limit]:
        print(f"{r['doc']}/{r['slide']:<7}{r['defect_nopic']:>9.4f}{r['defect']:>9.4f}"
              f"{r['open_area']:>8.1f}{r['mean_err']:>11.2f}{r['heavy']:>9.2f}")
    print(f"\nwrote {OUT}  ({len(rows)} slides)")


if __name__ == "__main__":
    main()
