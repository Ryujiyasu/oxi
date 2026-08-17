# -*- coding: utf-8 -*-
"""Census DrawingML features on the SLIDE-level shapes of the pptx dev corpus.

The ledger's existing census counted what the layout/master INHERITANCE gate
rejects. This one counts what the slides themselves carry, which is what the
renderer actually has to draw. Counts are (occurrences, slides, decks) per
feature so a single deck spamming one feature cannot look like corpus-wide
pressure.

Usage:
    python tools/metrics/pptx_feature_census.py [--decks d01,d04] [--top 40]
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = Path(__file__).resolve().parents[2]
PPTX_DIR = REPO_ROOT / "pipeline_data" / "pptx_benchmark" / "dev" / "pptx"

# Element-level features: counted by start tag.
ELEMENTS = [
    "a:custGeom", "a:prstGeom", "p:grpSp", "a:gradFill", "a:pattFill",
    "a:blipFill", "a:outerShdw", "a:innerShdw", "a:softEdge", "a:glow",
    "a:reflection", "p:cxnSp", "p:graphicFrame", "a:tbl", "p:pic", "a:effectLst",
    "a:alpha", "a:lumMod", "a:tint", "a:shade", "a:normAutofit", "a:spAutoFit",
    "a:bodyPr", "a:latin", "a:ea", "a:hlinkClick", "mc:AlternateContent",
]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--decks", default=None)
    parser.add_argument("--top", type=int, default=40)
    args = parser.parse_args()

    decks = sorted(PPTX_DIR.glob("*.pptx"))
    if args.decks:
        wanted = {s.strip() for s in args.decks.split(",")}
        decks = [d for d in decks if d.stem.split("__")[0] in wanted]

    occ: Counter[str] = Counter()
    slides_with: defaultdict[str, set] = defaultdict(set)
    decks_with: defaultdict[str, set] = defaultdict(set)
    prst: Counter[str] = Counter()
    prst_decks: defaultdict[str, set] = defaultdict(set)
    rot_occ = 0
    rot_slides: set = set()
    total_slides = 0

    for deck in decks:
        did = deck.stem.split("__")[0]
        with zipfile.ZipFile(deck) as zf:
            names = [n for n in zf.namelist()
                     if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)]
            total_slides += len(names)
            for name in names:
                key = f"{did}/{name.rsplit('/', 1)[1]}"
                xml = zf.read(name).decode("utf-8", errors="replace")
                for feature in ELEMENTS:
                    n = xml.count(f"<{feature}")
                    if n:
                        occ[feature] += n
                        slides_with[feature].add(key)
                        decks_with[feature].add(did)
                for m in re.finditer(r'<a:prstGeom[^>]*prst="([^"]+)"', xml):
                    prst[m.group(1)] += 1
                    prst_decks[m.group(1)].add(did)
                rots = [m for m in re.finditer(r'<a:xfrm[^>]*\brot="(-?\d+)"', xml)
                        if m.group(1) != "0"]
                if rots:
                    rot_occ += len(rots)
                    rot_slides.add(key)

    print(f"decks={len(decks)}  slides={total_slides}\n")
    print(f"{'feature':22s} {'occurrences':>11s} {'slides':>7s} {'decks':>6s}")
    for feature, n in occ.most_common():
        print(f"{feature:22s} {n:11d} {len(slides_with[feature]):7d} {len(decks_with[feature]):6d}")
    print(f"{'a:xfrm@rot!=0':22s} {rot_occ:11d} {len(rot_slides):7d}")

    print(f"\nprstGeom histogram (top {args.top})")
    for name, n in prst.most_common(args.top):
        print(f"  {name:24s} {n:7d}  {len(prst_decks[name]):2d} decks")


if __name__ == "__main__":
    main()
