# -*- coding: utf-8 -*-
"""How many dev-corpus paragraphs mix run styles inside one line?

`gdi_wrap_lines` measures a whole paragraph with ONE family / size / weight
(the first run's face, bold if ANY run is bold), so a paragraph that switches
style mid-sentence is measured wrong and breaks in the wrong place. This counts
the exposure: paragraphs with 2+ runs that differ in bold, italic, size or
typeface, and how much text sits in them.

Usage: python tools/metrics/pptx_mixedrun_census.py [--decks d19,d24]
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


def run_style(rpr: str) -> tuple:
    b = re.search(r'\bb="([01])"', rpr)
    i = re.search(r'\bi="([01])"', rpr)
    sz = re.search(r'\bsz="(\d+)"', rpr)
    lat = re.search(r'<a:latin typeface="([^"]+)"', rpr)
    return (
        b.group(1) if b else None,
        i.group(1) if i else None,
        sz.group(1) if sz else None,
        lat.group(1) if lat else None,
    )


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--decks", default=None)
    args = ap.parse_args()
    wanted = {s.strip() for s in args.decks.split(",")} if args.decks else None

    per_deck: Counter[str] = Counter()
    per_deck_total: Counter[str] = Counter()
    which: Counter[str] = Counter()
    for pptx in sorted(PPTX_DIR.glob("*.pptx")):
        did = pptx.stem.split("__")[0]
        if wanted and did not in wanted:
            continue
        z = zipfile.ZipFile(pptx)
        for name in z.namelist():
            if not re.match(r"ppt/slides/slide\d+\.xml$", name):
                continue
            xml = z.read(name).decode("utf-8", "replace")
            for para in re.finditer(r"<a:p>.*?</a:p>", xml, re.S):
                p = para.group(0)
                runs = re.findall(r"<a:r>(.*?)</a:r>", p, re.S)
                texts = [
                    "".join(re.findall(r"<a:t>([^<]*)</a:t>", r)) for r in runs
                ]
                if sum(len(t.strip()) for t in texts) == 0:
                    continue
                per_deck_total[did] += 1
                if len(runs) < 2:
                    continue
                styles = [run_style(r) for r in runs]
                if len(set(styles)) < 2:
                    continue
                per_deck[did] += 1
                for k, (a, b) in enumerate(zip(styles, styles[1:])):
                    for idx, label in enumerate(("bold", "italic", "size", "face")):
                        if a[idx] != b[idx]:
                            which[label] += 1
    total = sum(per_deck.values())
    grand = sum(per_deck_total.values())
    print(f"{total}/{grand} non-empty paragraphs mix run styles "
          f"({100 * total / max(grand, 1):.1f}%)")
    print("what differs:", dict(which))
    for did, n in per_deck.most_common(12):
        print(f"  {did}: {n}/{per_deck_total[did]}")


if __name__ == "__main__":
    main()
