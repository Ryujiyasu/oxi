# -*- coding: utf-8 -*-
"""Word-fidelity measurement for the India government-document corpus.

A corpus of Indian public-sector .docx, kept apart from the gated dev/blind
corpora because its point is DISCOVERY -- where Oxi's rendering of Indic script
(Devanagari and its kin) departs from Word -- not a merge gate. Drop .docx into

    pipeline_data/india_corpus/docx/

and run this. For each document it renders the Word truth (COM -> PDF -> PNG,
the same path pipeline.word_renderer uses) and the Oxi render (the DirectWrite
renderer, the pipeline default), scores every page with the production SSIM
(pipeline.ssim_calculator's own load/resize + skimage), and reports per-doc
mean and floor plus the corpus mean and worst page. A document Oxi cannot
render at all is reported as a rendering FAILURE, which for this corpus is the
most interesting result of all.

    python tools/metrics/india_corpus.py            # measure new/changed docs
    python tools/metrics/india_corpus.py --rerender # ignore cached PNGs
    python tools/metrics/india_corpus.py --docs foo,bar

Word COM makes this Windows-only, and -- like every Oxi render measurement --
nothing else may render while it runs (the renderers are not parallel-safe).
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

import numpy as np
from skimage.metrics import structural_similarity as ssim

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
from pipeline.config import RENDER_DPI  # noqa: E402
from pipeline.ssim_calculator import _load_rgb, _resize_to_match  # noqa: E402
from pipeline.word_renderer import _RENDER_SCRIPT  # noqa: E402  (Word COM -> PDF -> PNG)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = REPO / "pipeline_data" / "india_corpus"
DOCX_DIR = ROOT / "docx"
WORD_PNG = ROOT / "word_png"
OXI_PNG = ROOT / "oxi_png"
OUT = ROOT / "_scores.json"
WORD_RENDERER = REPO / "pipeline" / "word_renderer.py"
DWRITE = REPO / "tools" / "oxi-dwrite-renderer" / "target" / "release" / "oxi-dwrite-renderer.exe"
WORD_TIMEOUT = 60
OXI_TIMEOUT = 300


def render_word(docx: Path, out_dir: Path, rerender: bool) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    have = sorted(out_dir.glob("page_*.png"))
    if have and not rerender:
        return have
    for p in out_dir.glob("page_*.png"):
        p.unlink()
    # The pipeline renders one doc by running _RENDER_SCRIPT (Word COM -> PDF ->
    # PNG) in a child; do the same, aimed at this corpus's own output dir.
    r = subprocess.run(
        [sys.executable, "-c", _RENDER_SCRIPT, str(docx.resolve()), str(out_dir), str(RENDER_DPI)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=WORD_TIMEOUT, cwd=str(REPO),
    )
    if r.returncode != 0:
        print("   Word FAILED:", (r.stderr or r.stdout or "").strip()[:200])
    return sorted(out_dir.glob("page_*.png"))


def render_oxi(docx: Path, out_dir: Path, rerender: bool) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    have = sorted(out_dir.glob("page_*.png"))
    if have and not rerender:
        return have
    for p in out_dir.glob("*.png"):
        p.unlink()
    prefix = out_dir / "oxi"
    try:
        r = subprocess.run(
            [str(DWRITE), str(docx.resolve()), str(prefix), str(RENDER_DPI)],
            capture_output=True, timeout=OXI_TIMEOUT,
        )
    except subprocess.TimeoutExpired:
        print("   Oxi TIMEOUT")
        r = None
    if r is not None and r.returncode != 0:
        print("   Oxi FAILED:", r.stderr.decode("utf-8", "replace").strip()[:200])
    # Normalise oxi_pN.png -> page_NNNN.png
    i = 1
    while True:
        src = out_dir / f"oxi_p{i}.png"
        if not src.exists():
            break
        src.rename(out_dir / f"page_{i:04d}.png")
        i += 1
    return sorted(out_dir.glob("page_*.png"))


def score_pages(word_pngs: list[Path], oxi_pngs: list[Path]) -> tuple[list[float], int]:
    n = min(len(word_pngs), len(oxi_pngs))
    vals = []
    for i in range(n):
        # Word is the reference; the Oxi page is resized to it -- the same way
        # ssim_ab and the production gate compare.
        w = _load_rgb(str(word_pngs[i]))
        o = _resize_to_match(_load_rgb(str(oxi_pngs[i])), w)
        vals.append(float(ssim(w, o, channel_axis=2, data_range=255)))
    return vals, n


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--rerender", action="store_true")
    ap.add_argument("--docs", default="")
    args = ap.parse_args()

    docs = sorted(DOCX_DIR.glob("*.docx"))
    if args.docs:
        want = {d.strip() for d in args.docs.split(",") if d.strip()}
        docs = [d for d in docs if d.stem in want or any(w in d.stem for w in want)]
    if not docs:
        print(f"No .docx in {DOCX_DIR}")
        print("Drop Indian government .docx there, then re-run.")
        return

    state = {}
    if OUT.exists() and not args.rerender:
        try:
            state = json.loads(OUT.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            state = {}

    for docx in docs:
        doc_id = docx.stem
        print(doc_id)
        word_pngs = render_word(docx, WORD_PNG / doc_id, args.rerender)
        oxi_pngs = render_oxi(docx, OXI_PNG / doc_id, args.rerender)
        if not word_pngs:
            state[doc_id] = {"error": "word render produced no pages"}
        elif not oxi_pngs:
            state[doc_id] = {"error": "oxi render produced no pages",
                             "word_pages": len(word_pngs)}
        else:
            vals, n = score_pages(word_pngs, oxi_pngs)
            state[doc_id] = {
                "mean": float(np.mean(vals)), "min": float(np.min(vals)),
                "pages": n, "word_pages": len(word_pngs), "oxi_pages": len(oxi_pngs),
            }
            print("   mean %.4f  min %.4f  (%d/%d pages, oxi %d)"
                  % (state[doc_id]["mean"], state[doc_id]["min"], n,
                     len(word_pngs), len(oxi_pngs)))
        OUT.write_text(json.dumps(state, indent=1, ensure_ascii=False), encoding="utf-8")

    good = {k: v for k, v in state.items() if "mean" in v}
    errs = {k: v for k, v in state.items() if "error" in v}
    print("\n=== India corpus: %d docs (%d scored, %d render failures) ==="
          % (len(state), len(good), len(errs)))
    if good:
        means = sorted((v["mean"], k) for k, v in good.items())
        print("corpus mean-of-means %.4f  worst deck-mean %.4f (%s)"
              % (np.mean([m for m, _ in means]), means[0][0], means[0][1]))
        floor = min((v["min"], k) for k, v in good.items())
        print("floor page %.4f (%s)" % (floor[0], floor[1]))
    for k, v in errs.items():
        print("  FAILED %s: %s" % (k, v["error"]))


if __name__ == "__main__":
    main()
