# -*- coding: utf-8 -*-
"""Rank pptx slides by the part of the SSIM gap a BLUR cannot explain.

Ranking by raw SSIM -- by deck mean or by per-slide MIN -- puts photographic
slides on top, and they are not defects. The reference PDF re-compresses every
bitmap as JPEG, so a full-bleed photo carries grain Oxi cannot reproduce and
should not try to; the more faithfully the vector content is drawn, the more the
score is dominated by that grain (`pptx_reference_jpeg_grain`). d33 s10 is the
specimen: corpus MIN at 0.7673, yet mean|err| is 5.4/255, heavy(>40) is 0.33%,
and every text line sits within 0.96pt of PowerPoint's.

A defect survives a small blur; grain and antialiasing do not. So score each
slide twice and keep the blurred residual:

    defect = 1 - SSIM(blur(ref), blur(oxi))        r = 1.0px at 150dpi

and report it beside the raw gap, so a slide whose gap COLLAPSES under blur is
visibly a texture case rather than something to chase.

Writes `_defect_rank.json` next to the scores and prints the worst slides.
Resumable: an existing json is reused for slides whose PNG has not changed.

Usage:
    python tools/metrics/pptx_defect_rank.py [--limit N] [--doc 33]
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image, ImageFilter
from scipy.ndimage import uniform_filter

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
SS = REPO / "pipeline_data" / "pptx_benchmark" / "ssim_pptx"
CACHE = SS / "_defect_rank.json"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"
BLUR = 1.0


def ssim(a: np.ndarray, b: np.ndarray) -> float:
    a, b = a.mean(axis=2), b.mean(axis=2)
    c1, c2, w = (0.01 * 255) ** 2, (0.03 * 255) ** 2, 7
    m1, m2 = uniform_filter(a, w), uniform_filter(b, w)
    s11 = uniform_filter(a * a, w) - m1 * m1
    s22 = uniform_filter(b * b, w) - m2 * m2
    s12 = uniform_filter(a * b, w) - m1 * m2
    return float((((2 * m1 * m2 + c1) * (2 * s12 + c2))
                  / ((m1 ** 2 + m2 ** 2 + c1) * (s11 + s22 + c2))).mean())


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--doc", default="")
    args = ap.parse_args()

    # A PNG older than the binary was drawn by a renderer that no longer
    # exists. Ranking those silently is how 2026-08-27's first run put d44 s23
    # and s4 at the top of the corpus: both PNGs predated S-PICFILLBOX, the ship
    # derived to fix that very slide. Rank them, but never without saying so.
    exe_mtime = EXE.stat().st_mtime_ns if EXE.exists() else 0
    cache = json.loads(CACHE.read_text(encoding="utf-8")) if CACHE.exists() else {}
    docs = sorted(p.name for p in (SS / "oxi_png").iterdir() if p.is_dir())
    if args.doc:
        docs = [d for d in docs if d == args.doc]
    rows = []
    for doc in docs:
        pdf_path = SS / "ppt_pdf" / f"{doc}.pdf"
        if not pdf_path.exists():
            continue
        pdf = pymupdf.open(pdf_path)
        for png in sorted((SS / "oxi_png" / doc).glob("slide_s*.png"),
                          key=lambda p: int(p.stem.split("_s")[1])):
            idx = int(png.stem.split("_s")[1])
            if idx > len(pdf):
                continue
            key = f"{doc}/{idx}"
            st = png.stat().st_mtime_ns
            hit = cache.get(key)
            if hit and hit.get("mtime") == st:
                rows.append(hit)
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
            row = {
                "doc": doc, "slide": idx, "mtime": st,
                "ssim": round(raw, 5),
                "ssim_blur": round(ssim(rb, ob), 5),
                "mean_err": round(float(np.abs(r - o).mean()), 3),
                "heavy": round(float((np.abs(r - o).mean(axis=2) > 40).mean() * 100), 3),
            }
            row["defect"] = round(1 - row["ssim_blur"], 5)
            row["explained"] = round((row["ssim_blur"] - raw) / max(1 - raw, 1e-6), 4)
            cache[key] = row
            rows.append(row)
        pdf.close()
        CACHE.write_text(json.dumps(cache, indent=0), encoding="utf-8")
        print(f"  {doc} done ({len(rows)} slides)", flush=True)

    rows.sort(key=lambda r: -r["defect"])
    # Recomputed every run rather than stored: a rebuild must make the whole
    # table go stale immediately, without anyone remembering to invalidate it.
    for r in rows:
        r["stale"] = r.get("mtime", 0) < exe_mtime
    stale = [r for r in rows if r["stale"]]
    if stale:
        docs = sorted({r["doc"] for r in stale})
        print()
        print(f"!! {len(stale)} of {len(rows)} slides were rendered BEFORE the "
              f"current binary and are marked * -- re-render before trusting them.")
        print(f"   decks affected: {','.join(docs)}")
        print("   fix: python tools/metrics/pptx_remeasure.py --force")
    n = args.limit or 30
    print(f"\nworst {n} slides by BLURRED residual (blur r={BLUR}px)")
    print(f"{'slide':<10}{'defect':>9}{'ssim':>9}{'blurred':>9}"
          f"{'explained':>11}{'mean|err|':>11}{'heavy%':>8}")
    for r in rows[:n]:
        tag = "*" if r["stale"] else " "
        print(f"{r['doc']}/s{r['slide']}{tag:<7}{r['defect']:>9.4f}{r['ssim']:>9.4f}"
              f"{r['ssim_blur']:>9.4f}{r['explained'] * 100:>10.1f}%"
              f"{r['mean_err']:>11.2f}{r['heavy']:>8.2f}")
    print(f"\n{len(rows)} slides scored -> {CACHE}")


if __name__ == "__main__":
    main()
