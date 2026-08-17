# -*- coding: utf-8 -*-
"""Byte-identity A/B over the pptx PROBE corpus (pipeline_data/pptx_probes).

The probe decks pin chart geometry, bullets, backgrounds and theme resolution.
A change that is supposed to touch only one feature must leave every probe deck
BYTE-IDENTICAL -- which implies their SSIM batteries are unchanged without
having to re-score them.

Renders each deck twice with the SAME binary, once per env arm, and compares
the PNGs. Note `pptx_probes` nests one level deeper than you expect, so the
glob must be recursive (a past session's gate silently scored nothing).

Usage:
    python tools/metrics/pptx_probe_ab.py --env OXI_CUSTGEOM_DISABLE=1
"""
from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = Path(__file__).resolve().parents[2]
PROBES = REPO_ROOT / "pipeline_data" / "pptx_probes"
OXI_EXE = REPO_ROOT / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"


def render(pptx: Path, out_dir: Path, env_pairs: dict[str, str]) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ)
    env.update(env_pairs)
    subprocess.run(
        [str(OXI_EXE), str(pptx), str(out_dir / "slide"), "150"],
        capture_output=True,
        timeout=900,
        env=env,
    )
    return sorted(out_dir.glob("slide_s*.png"), key=lambda p: int(p.stem.split("_s")[1]))


def one(pptx: Path, env_pairs: dict[str, str], root: Path) -> tuple[str, int, int, str]:
    stem = pptx.parent.name + "/" + pptx.stem
    base = root / hashlib.md5(str(pptx).encode()).hexdigest()[:10]
    a = render(pptx, base / "a", {})
    b = render(pptx, base / "b", env_pairs)
    if len(a) != len(b):
        return stem, len(a), len(b), "PAGE COUNT"
    diffs = sum(
        1 for pa, pb in zip(a, b)
        if hashlib.sha256(pa.read_bytes()).digest() != hashlib.sha256(pb.read_bytes()).digest()
    )
    shutil.rmtree(base, ignore_errors=True)
    return stem, len(a), diffs, "CHANGED" if diffs else "same"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--env", action="append", default=[], required=True)
    ap.add_argument("--jobs", type=int, default=5)
    args = ap.parse_args()
    env_pairs = dict(pair.split("=", 1) for pair in args.env)

    decks = sorted(PROBES.rglob("*.pptx"))
    print(f"probe decks: {len(decks)}   arm B env: {env_pairs}")
    root = Path(tempfile.mkdtemp(prefix="pptx_probe_ab_"))
    pages = changed = same = 0
    try:
        with ThreadPoolExecutor(max_workers=args.jobs) as pool:
            futures = [pool.submit(one, d, env_pairs, root) for d in decks]
            for future in as_completed(futures):
                stem, n, diffs, verdict = future.result()
                pages += n
                if verdict == "same":
                    same += 1
                else:
                    changed += 1
                    print(f"  {verdict} {stem}: {diffs}/{n}")
    finally:
        shutil.rmtree(root, ignore_errors=True)
    print(f"\n{len(decks)} decks / {pages} pages: byte-identical {same} / CHANGED {changed}")


if __name__ == "__main__":
    main()
