# -*- coding: utf-8 -*-
"""A/B one opt-out flag over the blind pptx corpus, with the SAME binary.

The control is the patched binary run with the feature's `OXI_*_DISABLE` set,
never a cached render and never the previous build: a cache stops being a valid
control at the next ship (2026-08-27, a stale cache credited three commits to
one and showed a false +0.0215 on d31), and a different build changes more than
the flag does.

Renders each deck twice, arm OFF then arm ON, scores both against PowerPoint's
PDF, and reports per-deck mean and MIN plus the slides that moved. A deck whose
two arms are byte-identical is reported as untouched, which is the arm-A proof
for every deck the feature is not supposed to reach.

★The renderer is NOT parallel-safe (`pptx_render_not_parallel_safe`): one deck,
one arm, at a time -- and nothing else may render while this runs.

Usage:
    python tools/metrics/pptx_flag_ab.py OXI_CELLADV_DISABLE [--docs 33,03,23]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image
from skimage.metrics import structural_similarity as ssim

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
SSIMD = ROOT / "ssim_pptx"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"
DPI = 150


def render(src: Path, out: Path, flag: str, on: bool) -> list[Path]:
    """`on` selects the FEATURE, not the variable: an `_ENABLE` flag is set to
    turn the feature on, a `_DISABLE` flag is set to turn it off. Passing the
    variable through unchanged from the environment would let an inherited value
    decide the arm, so it is always removed first."""
    shutil.rmtree(out, ignore_errors=True)
    out.mkdir(parents=True)
    env = dict(os.environ)
    env.pop(flag, None)
    want_set = (not on) if flag.endswith("_DISABLE") else on
    if want_set:
        env[flag] = "1"
    r = subprocess.run([str(EXE), str(src), str(out / "slide"), str(DPI)],
                       capture_output=True, env=env, timeout=3600)
    if r.returncode != 0:
        return []
    return sorted(out.glob("slide_s*.png"), key=lambda p: int(p.stem.split("_s")[1]))


def score(pdf, pngs: list[Path]) -> list[float]:
    vals = []
    for i in range(len(pdf)):
        p = pngs[i] if i < len(pngs) else None
        if p is None or not p.exists():
            continue
        pix = pdf[i].get_pixmap(dpi=DPI)
        a = np.asarray(Image.frombytes("RGB", (pix.width, pix.height), pix.samples)).astype(float)
        o = Image.open(p).convert("RGB")
        if o.size != (pix.width, pix.height):
            o = o.resize((pix.width, pix.height), Image.LANCZOS)
        vals.append(ssim(a, np.asarray(o).astype(float), channel_axis=2, data_range=255))
    return vals


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("flag")
    ap.add_argument("--docs", default="")
    args = ap.parse_args()

    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    want = {d.strip().lstrip("0") or "0" for d in args.docs.split(",") if d.strip()}
    tmp_a, tmp_b = SSIMD / "_ab_off", SSIMD / "_ab_on"
    tot_a = tot_b = 0.0
    n_deck = up = down = untouched = 0
    for item in manifest:
        doc = f"{item['idx']:02d}"
        if want and str(item["idx"]) not in want:
            continue
        src = ROOT / "pptx" / item["local"]
        pdf_path = SSIMD / "ppt_pdf" / f"{doc}.pdf"
        if not src.exists() or not pdf_path.exists():
            continue
        t0 = time.time()
        pdf = pymupdf.open(pdf_path)
        pa = render(src, tmp_a, args.flag, on=False)   # control: feature OFF
        pb = render(src, tmp_b, args.flag, on=True)    # treatment: feature ON
        if not pa or not pb:
            print(f"{doc}: render failed", flush=True)
            pdf.close()
            continue
        ha = [hashlib.sha256(p.read_bytes()).hexdigest() for p in pa]
        hb = [hashlib.sha256(p.read_bytes()).hexdigest() for p in pb]
        if ha == hb:
            untouched += 1
            print(f"{doc}: untouched (byte-identical, {len(pa)} slides) [{time.time()-t0:.0f}s]",
                  flush=True)
            pdf.close()
            continue
        va, vb = score(pdf, pa), score(pdf, pb)
        pdf.close()
        n = min(len(va), len(vb))
        if not n:
            continue
        ma, mb = sum(va[:n]) / n, sum(vb[:n]) / n
        tot_a += ma
        tot_b += mb
        n_deck += 1
        moved = [(i + 1, vb[i] - va[i]) for i in range(n) if abs(vb[i] - va[i]) > 1e-6]
        up += sum(1 for _, d in moved if d > 0)
        down += sum(1 for _, d in moved if d < 0)
        print(f"{doc}: {ma:.6f} -> {mb:.6f} ({mb-ma:+.6f})  "
              f"MIN {min(va[:n]):.4f} -> {min(vb[:n]):.4f}  "
              f"moved {len(moved)}/{n} [{time.time()-t0:.0f}s]", flush=True)
        for s, d in sorted(moved, key=lambda x: -abs(x[1]))[:6]:
            print(f"    s{s} {d:+.4f}")
    shutil.rmtree(tmp_a, ignore_errors=True)
    shutil.rmtree(tmp_b, ignore_errors=True)
    if n_deck:
        print(f"\n{n_deck} decks moved ({untouched} byte-identical): "
              f"{tot_a/n_deck:.6f} -> {tot_b/n_deck:.6f} ({(tot_b-tot_a)/n_deck:+.6f})")
        print(f"slides up {up}  down {down}")
    else:
        print(f"\nno deck moved ({untouched} byte-identical)")


if __name__ == "__main__":
    main()
