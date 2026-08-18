# -*- coding: utf-8 -*-
"""Does the renderer produce the same bytes twice for the same input?

Every byte-identity gate in the pptx workflow -- the arm-A proof, the probe
corpus check -- assumes it does. It did not: on 2026-08-18 the same binary
rendered d19 slide 39 (an icon row) shifted by one glyph on a second run,
because `GetCharABCWidthsW` consults GDI's font-link chain and its success is
not stable, so the design-advance draw path was taken on one run and not the
other. The layout dump was byte-identical both times, i.e. nothing upstream of
drawing varied.

Run this after any change that touches font selection or measurement, and
before trusting a byte-identity result.

Usage:
    python tools/metrics/pptx_determinism.py                 # a sample of decks
    python tools/metrics/pptx_determinism.py --all
    python tools/metrics/pptx_determinism.py --decks d19,d28 --runs 3
"""
from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = Path(__file__).resolve().parents[2]
DEV = REPO_ROOT / "pipeline_data" / "pptx_benchmark" / "dev" / "pptx"
EXE = REPO_ROOT / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"


def digest_dir(path: Path) -> dict[str, str]:
    return {
        p.name: hashlib.sha256(p.read_bytes()).hexdigest()
        for p in sorted(path.glob("slide_s*.png"))
    }


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--decks", default=None)
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--runs", type=int, default=2)
    args = ap.parse_args()

    decks = sorted(DEV.glob("*.pptx"))
    if args.decks:
        wanted = {s.strip() for s in args.decks.split(",")}
        decks = [d for d in decks if d.stem.split("__")[0] in wanted]
    elif not args.all:
        # A cheap default: the decks that exercise icon rows, embedded fonts
        # and heavy custom geometry.
        wanted = {"d19", "d28", "d32", "d10"}
        decks = [d for d in decks if d.stem.split("__")[0] in wanted]

    root = Path(tempfile.mkdtemp(prefix="pptx_det_"))
    unstable = 0
    try:
        for deck in decks:
            digests = []
            for run in range(args.runs):
                out = root / f"{deck.stem}_{run}"
                out.mkdir(parents=True, exist_ok=True)
                subprocess.run(
                    [str(EXE), str(deck), str(out / "slide"), "150"],
                    capture_output=True,
                    timeout=1800,
                    env=dict(os.environ),
                )
                digests.append(digest_dir(out))
            base = digests[0]
            bad = sorted(
                {
                    name
                    for other in digests[1:]
                    for name in set(base) | set(other)
                    if base.get(name) != other.get(name)
                }
            )
            if bad:
                unstable += 1
                print(f"  UNSTABLE {deck.stem[:44]}: {len(bad)} page(s) {bad[:6]}")
            else:
                print(f"  stable   {deck.stem[:44]}: {len(base)} pages x {args.runs} runs")
    finally:
        shutil.rmtree(root, ignore_errors=True)
    print(f"\n{len(decks)} decks / {args.runs} runs each: {unstable} unstable")
    sys.exit(1 if unstable else 0)


if __name__ == "__main__":
    main()
