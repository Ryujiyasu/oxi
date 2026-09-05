# -*- coding: utf-8 -*-
"""The blind corpus's absolute score with the CURRENT binary, one arm.

Renders every blind deck once at the A/B tool's DPI and scores each slide
against the stored truth PDF exactly the way `pptx_flag_ab.py` does (gray
SSIM at 150 DPI), then prints per-deck means, the corpus mean and the floor.
The A/B tool answers "did this flag move things"; this answers "where does
the corpus stand".

★The renderer is not parallel-safe: decks render one at a time, and nothing
else may render while this runs.

    python tools/metrics/pptx_blind_score.py [--docs 01,02]
"""
from __future__ import annotations

import argparse
import json
import os
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
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"
OUT = ROOT / "ssim_pptx" / "_blind_score.json"
DPI = 150


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docs", default="")
    args = ap.parse_args()
    man = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    docs = [f"{m['idx']:02d}" for m in man]
    if args.docs:
        want = {d.strip().zfill(2) for d in args.docs.split(",") if d.strip()}
        docs = [d for d in docs if d in want]

    rows = {}
    for doc in docs:
        hits = sorted((ROOT / "pptx").glob(doc + "*.pptx"))
        pdfs = sorted((ROOT / "ssim_pptx" / "ppt_pdf").glob(doc + "*.pdf"))
        if not hits or not pdfs:
            print(doc, "missing inputs")
            continue
        outdir = ROOT / "ssim_pptx" / "_blind_score_png" / doc
        outdir.mkdir(parents=True, exist_ok=True)
        for f in outdir.glob("*.png"):
            f.unlink()
        r = subprocess.run([str(EXE), str(hits[0]), str(outdir / "slide"), str(DPI)],
                           capture_output=True)
        if r.returncode != 0:
            print(doc, "render failed", r.returncode)
            continue
        pdf = pymupdf.open(pdfs[0])
        scores = []
        for i in range(pdf.page_count):
            png = outdir / f"slide_s{i + 1}.png"
            if not png.exists():
                continue
            oxi = Image.open(png).convert("L")
            pix = pdf[i].get_pixmap(dpi=DPI)
            ref = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).convert("L")
            if ref.size != oxi.size:
                ref = ref.resize(oxi.size)
            scores.append(ssim(np.asarray(ref, np.float64),
                               np.asarray(oxi, np.float64), data_range=255))
        pdf.close()
        if scores:
            rows[doc] = {"mean": float(np.mean(scores)), "min": float(np.min(scores)),
                         "n": len(scores)}
            print("%s  mean %.6f  min %.4f  (%d slides)"
                  % (doc, rows[doc]["mean"], rows[doc]["min"], len(scores)))
    if rows:
        means = [r["mean"] for r in rows.values()]
        mins = [(r["min"], d) for d, r in rows.items()]
        print("\ncorpus: %d decks  mean-of-means %.6f  worst deck-mean %.6f (%s)  "
              "floor slide %.4f (%s)"
              % (len(rows), np.mean(means), min(means),
                 min(rows, key=lambda d: rows[d]["mean"]),
                 min(mins)[0], min(mins)[1]))
        OUT.write_text(json.dumps(rows, indent=1), encoding="utf-8")
        print("wrote", OUT)


if __name__ == "__main__":
    main()
