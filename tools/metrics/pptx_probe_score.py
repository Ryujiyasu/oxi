# -*- coding: utf-8 -*-
"""Score probe decks against PowerPoint's own PDF, one arm against another.

`pptx_probe_ab.py` answers "did this change touch the probes at all", which is
the right question for a change that must be inert. When a change is SUPPOSED to
move them -- a layout rule the probes were built to pin -- the question becomes
"did they move toward PowerPoint", and that needs the truth PDF each probe
directory already carries.

Usage:
    python tools/metrics/pptx_probe_score.py --env OXI_MIXPITCH_DISABLE=1 \
        --decks spec4b_lspacing,spec4c_lspacing
Arm A is the plain environment (the change ON), arm B adds --env.
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
PROBES = REPO / "pipeline_data" / "pptx_probes"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"
DPI = 150


def find_pdf(deck: Path) -> Path | None:
    """The truth PDF sits next to the deck, or in a sibling `<prefix>_truth`."""
    if deck.with_suffix(".pdf").exists():
        return deck.with_suffix(".pdf")
    hits = sorted(PROBES.rglob(f"{deck.stem}.pdf"))
    return hits[0] if hits else None


def ssim(a: np.ndarray, b: np.ndarray) -> float:
    """Global SSIM on the luma plane -- enough to rank two arms."""
    a, b = a.astype(float), b.astype(float)
    c1, c2 = (0.01 * 255) ** 2, (0.03 * 255) ** 2
    ma, mb, va, vb = a.mean(), b.mean(), a.var(), b.var()
    cov = ((a - ma) * (b - mb)).mean()
    return ((2 * ma * mb + c1) * (2 * cov + c2)) / ((ma * ma + mb * mb + c1) * (va + vb + c2))


def render(deck: Path, out: Path, env: dict[str, str]) -> list[Path]:
    out.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [str(EXE), str(deck), str(out / "slide"), str(DPI)],
        capture_output=True, timeout=1800, env={**os.environ, **env},
    )
    return sorted(out.glob("slide_s*.png"), key=lambda p: int(p.stem.split("_s")[1]))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--env", required=True, help="VAR=VALUE applied to arm B")
    ap.add_argument("--decks", required=True, help="comma-separated probe stems")
    ap.add_argument("--pdf", default=None,
                    help="truth PDF to use for a single deck whose file is named "
                         "something else (e.g. truth_geometry/deck.pdf)")
    args = ap.parse_args()
    k, _, v = args.env.partition("=")
    root = Path(tempfile.mkdtemp(prefix="probe_score_"))
    try:
        for stem in args.decks.split(","):
            hits = list(PROBES.rglob(f"{stem.strip()}.pptx"))
            if not hits:
                print(f"{stem}: no such probe")
                continue
            deck = hits[0]
            pdf = Path(args.pdf).resolve() if args.pdf else find_pdf(deck)
            if pdf is None:
                print(f"{deck.stem}: no truth PDF next to it")
                continue
            doc = pymupdf.open(pdf)
            a = render(deck, root / f"{deck.stem}_a", {})
            b = render(deck, root / f"{deck.stem}_b", {k: v})
            print(f"\n{deck.stem}  ({pdf.relative_to(PROBES)})")
            tot_a = tot_b = 0.0
            n = 0
            for i, (pa, pb) in enumerate(zip(a, b)):
                if i >= doc.page_count:
                    break
                ref = doc[i].get_pixmap(dpi=DPI)
                ref_img = np.asarray(
                    Image.frombytes("RGB", (ref.width, ref.height), ref.samples).convert("L")
                )
                ia = np.asarray(Image.open(pa).convert("L"))
                ib = np.asarray(Image.open(pb).convert("L"))
                h = min(ref_img.shape[0], ia.shape[0], ib.shape[0])
                w = min(ref_img.shape[1], ia.shape[1], ib.shape[1])
                sa = ssim(ref_img[:h, :w], ia[:h, :w])
                sb = ssim(ref_img[:h, :w], ib[:h, :w])
                tot_a += sa; tot_b += sb; n += 1
                mark = "  " if abs(sa - sb) < 1e-6 else ("<-" if sa > sb else "->")
                print(f"   p{i + 1:<3d} on={sa:.4f}  off={sb:.4f}  {mark}")
            if n:
                print(f"   MEAN on={tot_a / n:.4f}  off={tot_b / n:.4f}  "
                      f"delta={(tot_a - tot_b) / n:+.4f}")
    finally:
        shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    main()
