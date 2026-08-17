# -*- coding: utf-8 -*-
"""Byte-identity A/B between two rendered arms of the pptx dev corpus.

The ledger's arm-A proof: render the corpus with the NEW binary and the
feature's opt-out set, then show it is byte-identical to the pre-patch
binary's output. That is what makes the following SSIM A/B a real
before-vs-after rather than a comparison of two different builds.

Usage:
    python tools/metrics/pptx_byte_ab.py head off
"""
from __future__ import annotations

import hashlib
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PNG_ROOT = REPO_ROOT / "pipeline_data" / "pptx_benchmark" / "dev" / "oxi_png"


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> None:
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    a_tag, b_tag = sys.argv[1], sys.argv[2]
    a_root, b_root = PNG_ROOT / a_tag, PNG_ROOT / b_tag
    decks = sorted({p.name for p in a_root.iterdir()} | {p.name for p in b_root.iterdir()})
    same_decks = changed = missing = 0
    changed_pages = 0
    for deck in decks:
        a_pngs = {p.name: p for p in (a_root / deck).glob("slide_s*.png")} if (a_root / deck).is_dir() else {}
        b_pngs = {p.name: p for p in (b_root / deck).glob("slide_s*.png")} if (b_root / deck).is_dir() else {}
        if a_pngs.keys() != b_pngs.keys():
            print(f"  PAGE-SET DIFFERS {deck}: {len(a_pngs)} vs {len(b_pngs)}")
            missing += 1
            continue
        diffs = [n for n in a_pngs if digest(a_pngs[n]) != digest(b_pngs[n])]
        if diffs:
            changed += 1
            changed_pages += len(diffs)
            print(f"  CHANGED {deck}: {len(diffs)}/{len(a_pngs)} pages")
        else:
            same_decks += 1
    total = same_decks + changed + missing
    print(
        f"\n{a_tag} vs {b_tag}: {total} decks / byte-identical {same_decks} / "
        f"CHANGED {changed} ({changed_pages} pages) / page-set differs {missing}"
    )


if __name__ == "__main__":
    main()
