# -*- coding: utf-8 -*-
"""How much of the dev corpus uses a:prstTxWarp (WordArt text fitting)?

A shape whose bodyPr carries `<a:prstTxWarp>` has its text scaled to the shape
box by PowerPoint — d35's transition slides put a single numeral in a
75.9 x 303.5pt box with no `sz` anywhere and PowerPoint draws it ~300pt tall,
while Oxi falls back to its default size and draws a small character in the
corner. This counts the exposure and reports the box each one sits in.

Usage: python tools/metrics/pptx_txwarp_census.py [--decks d35,d11]
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
PPTX_DIR = REPO / "pipeline_data" / "pptx_benchmark" / "dev" / "pptx"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--decks", default=None)
    ap.add_argument("--show", type=int, default=10)
    args = ap.parse_args()
    wanted = {s.strip() for s in args.decks.split(",")} if args.decks else None

    presets: Counter[str] = Counter()
    per_deck: Counter[str] = Counter()
    sized = unsized = 0
    examples = []
    for pptx in sorted(PPTX_DIR.glob("*.pptx")):
        did = pptx.stem.split("__")[0]
        if wanted and did not in wanted:
            continue
        z = zipfile.ZipFile(pptx)
        for name in z.namelist():
            m = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
            if not m:
                continue
            xml = z.read(name).decode("utf-8", "replace")
            if "prstTxWarp" not in xml:
                continue
            for sp in re.finditer(r"<p:sp>.*?</p:sp>", xml, re.S):
                blk = sp.group(0)
                warp = re.search(r'<a:prstTxWarp prst="([^"]+)"', blk)
                if not warp:
                    continue
                presets[warp.group(1)] += 1
                per_deck[did] += 1
                ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', blk)
                has_sz = bool(re.search(r'sz="\d+"', blk))
                if has_sz:
                    sized += 1
                else:
                    unsized += 1
                text = "".join(re.findall(r"<a:t>([^<]*)</a:t>", blk))[:18]
                if len(examples) < args.show:
                    box = (tuple(round(int(v) / 12700, 1) for v in ext.groups())
                           if ext else None)
                    examples.append(
                        f"{did} s{m.group(1)} {warp.group(1)} box={box} "
                        f"sz={'yes' if has_sz else 'NONE'} {text!r}")
    print(f"presets: {dict(presets)}")
    print(f"{sized} with an explicit size, {unsized} WITHOUT")
    for did, n in per_deck.most_common(12):
        print(f"  {did}: {n}")
    print()
    for e in examples:
        print("  " + e)


if __name__ == "__main__":
    main()
