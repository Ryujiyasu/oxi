# -*- coding: utf-8 -*-
"""Score the blind pptx set in parallel, because the sequential one need not be.

`pptx_blind_score.py` walks the corpus a deck at a time and takes the better
part of an hour. Nothing in it touches the renderer or PowerPoint -- it reads
two PNG-sized arrays and runs SSIM -- so the only reason it was serial is that
it was written beside a tool that had to be (`pptx_render_not_parallel_safe`
governs the RENDERER, not the scorer).

Same arithmetic, same output file, one process per deck.

    python tools/metrics/pptx_blind_score_par.py [--jobs N]
"""
from __future__ import annotations

import argparse
import json
import sys
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image
from skimage.metrics import structural_similarity as ssim

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
SS = ROOT / "ssim_pptx"
OUT = SS / "_incremental.json"
DPI = 150

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def score_deck(doc: str) -> tuple[str, dict | None]:
    pdf_path = SS / "ppt_pdf" / f"{doc}.pdf"
    png_dir = SS / "oxi_png" / doc
    if not pdf_path.exists() or not png_dir.is_dir():
        return doc, None
    pdf = pymupdf.open(pdf_path)
    pngs = sorted(png_dir.glob("slide_s*.png"),
                  key=lambda p: int(p.stem.split("_s")[1]))
    # ★The partial-cache trap the sequential tool documents: a deck whose cache
    # is short of the truth's page count is refused, not quietly scored on the
    # slides that happen to be there.
    if len(pngs) < pdf.page_count:
        pdf.close()
        return doc, {"error": f"cache has {len(pngs)} of {pdf.page_count} slides"}
    vals = []
    for i in range(pdf.page_count):
        pix = pdf[i].get_pixmap(dpi=DPI)
        a = np.asarray(Image.frombytes("RGB", (pix.width, pix.height), pix.samples)).astype(float)
        o = Image.open(pngs[i]).convert("RGB")
        if o.size != (pix.width, pix.height):
            o = o.resize((pix.width, pix.height), Image.LANCZOS)
        vals.append(float(ssim(a, np.asarray(o).astype(float), channel_axis=2, data_range=255)))
    pdf.close()
    return doc, {"mean": sum(vals) / len(vals), "min": min(vals), "slides": vals}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--jobs", type=int, default=0)
    args = ap.parse_args()
    man = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    docs = [f"{m['idx']:02d}" for m in man]
    import os
    jobs = args.jobs or max(1, (os.cpu_count() or 4) - 1)
    print(f"{len(docs)} decks, {jobs} at a time\n", flush=True)
    got: dict[str, dict] = {}
    with ProcessPoolExecutor(max_workers=jobs) as pool:
        for doc, res in pool.map(score_deck, docs):
            if res is None:
                continue
            if "error" in res:
                print(f"{doc}: {res['error']}", flush=True)
                continue
            got[doc] = res
            print(f"{doc}: mean {res['mean']:.4f}  min {res['min']:.4f} "
                  f"({len(res['slides'])} slides)", flush=True)
    OUT.write_text(json.dumps(got, indent=1), encoding="utf-8")
    if got:
        means = [v["mean"] for v in got.values()]
        worst = min(got.items(), key=lambda kv: kv[1]["mean"])
        floor = min(got.items(), key=lambda kv: kv[1]["min"])
        print(f"\n{len(got)} decks: mean of means {sum(means) / len(means):.4f}")
        print(f"worst deck {worst[0]} at {worst[1]['mean']:.4f}; "
              f"lowest single slide {floor[1]['min']:.4f} in deck {floor[0]}")
        print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
