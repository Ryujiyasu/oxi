# -*- coding: utf-8 -*-
"""Score the blind pptx set one deck at a time, saving as it goes.

Two hard-won behaviours, both kept:

- **Incremental** (2026-08): the bundled harness writes its result once, at
  the end, and three interrupted runs in a row each discarded ~40 minutes of
  SSIM. This writes `_blind_score.json` after every deck, so a run can be
  stopped, resumed, or split across invocations.
- **Fresh renders** (2026-09-05): scoring a pre-filled cache trusts its
  freshness, which is the stale-binary trap. By default each deck is rendered
  with the CURRENT binary before scoring (one at a time -- the renderer is
  not parallel-safe, and nothing else may render while this runs).
  `--cache` restores scoring an existing `ssim_pptx/oxi_png/<doc>` cache, and
  then a deck whose cache is short of the PDF's slide count is REFUSED rather
  than quietly scored on the slides that happen to exist (the partial-cache
  trap that made a 19-slide deck read as a 4-slide one).

The arithmetic matches `pptx_flag_ab.py`: gray SSIM at 150 DPI against the
stored truth PDF, per slide, deck mean and min.

Usage:
    python tools/metrics/pptx_blind_score.py [--docs 01,02] [--cache]
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image
from skimage.metrics import structural_similarity as ssim

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
SSIM_DIR = ROOT / "ssim_pptx"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"
OUT = SSIM_DIR / "_blind_score.json"
DPI = 150


def load() -> dict:
    if OUT.exists():
        try:
            return json.loads(OUT.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save(state: dict) -> None:
    OUT.write_text(json.dumps(state, indent=1, ensure_ascii=False), encoding="utf-8")


def score_deck(doc: str, cache: bool) -> dict | None:
    pdfs = sorted((SSIM_DIR / "ppt_pdf").glob(doc + "*.pdf"))
    if not pdfs:
        return None
    if cache:
        png_dir = SSIM_DIR / "oxi_png" / doc
        if not png_dir.is_dir():
            return None
    else:
        hits = sorted((ROOT / "pptx").glob(doc + "*.pptx"))
        if not hits:
            return None
        png_dir = SSIM_DIR / "_blind_score_png" / doc
        png_dir.mkdir(parents=True, exist_ok=True)
        for f in png_dir.glob("*.png"):
            f.unlink()
        r = subprocess.run([str(EXE), str(hits[0]), str(png_dir / "slide"), str(DPI)],
                           capture_output=True)
        if r.returncode != 0:
            return {"doc": doc, "error": f"render failed rc={r.returncode}"}
    pdf = pymupdf.open(pdfs[0])
    have = list(png_dir.glob("slide_s*.png"))
    if len(have) < pdf.page_count:
        pdf.close()
        # A short cache is a lie waiting to be averaged.
        return {"doc": doc, "error": f"cache has {len(have)} of {pdf.page_count} slides"}
    vals = []
    for i in range(pdf.page_count):
        p = png_dir / f"slide_s{i + 1}.png"
        if not p.exists():
            continue
        pix = pdf[i].get_pixmap(dpi=DPI)
        ref = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).convert("L")
        oxi = Image.open(p).convert("L")
        if oxi.size != ref.size:
            ref = ref.resize(oxi.size)
        vals.append(ssim(np.asarray(ref, np.float64), np.asarray(oxi, np.float64),
                         data_range=255))
    pdf.close()
    if not vals:
        return {"doc": doc, "error": "no comparable slides"}
    return {"doc": doc, "mean": float(np.mean(vals)), "min": float(np.min(vals)),
            "n": len(vals)}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docs", default="")
    ap.add_argument("--cache", action="store_true",
                    help="score the existing oxi_png cache instead of rendering")
    args = ap.parse_args()
    man = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    docs = [f"{m['idx']:02d}" for m in man]
    if args.docs:
        want = {d.strip().zfill(2) for d in args.docs.split(",") if d.strip()}
        docs = [d for d in docs if d in want]

    state = load()
    for doc in docs:
        got = score_deck(doc, args.cache)
        if got is None:
            print(doc, "missing inputs")
            continue
        state[doc] = got
        save(state)
        if "error" in got:
            print(doc, got["error"])
        else:
            print("%s  mean %.6f  min %.4f  (%d slides)"
                  % (doc, got["mean"], got["min"], got["n"]))
    good = {d: r for d, r in state.items() if "mean" in r}
    if good:
        means = [r["mean"] for r in good.values()]
        worst = min(good, key=lambda d: good[d]["mean"])
        floor = min(good, key=lambda d: good[d]["min"])
        print("\ncorpus: %d decks  mean-of-means %.6f  worst deck-mean %.6f (%s)  "
              "floor slide %.4f (%s)"
              % (len(good), float(np.mean(means)), good[worst]["mean"], worst,
                 good[floor]["min"], floor))


if __name__ == "__main__":
    main()
