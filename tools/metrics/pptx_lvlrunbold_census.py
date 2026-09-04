# -*- coding: utf-8 -*-
"""Which decks can S-LVLRUNBOLD change?

The rule only bites where the per-run measurement runs at all and a run has no
weight of its own to measure with:

  1. the paragraph has TWO OR MORE runs -- a one-run paragraph is measured with
     the whole-line style, which already carried the level's weight;
  2. at least one of those runs declares no `b`;
  3. the placeholder's level says bold, so what the silent run inherits is not
     what it inherited before.

Condition 3 needs the slide -> layout -> master chain, so this over-approximates
it: any `b="1"` in a `defRPr` the deck's layouts or master carry. A deck the
census does not name renders identically under the flag; the A/B then confirms
that as byte-identical arms, at two renders a deck instead of a zip read.

    python tools/metrics/pptx_lvlrunbold_census.py [--blind]
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[2] / "pipeline_data" / "pptx_benchmark"
PARA = re.compile(r"<a:p>(.*?)</a:p>", re.S)
RUN = re.compile(r"<a:r>(.*?)</a:r>", re.S)
RPR = re.compile(r"<a:rPr\b([^>]*)>|<a:rPr\b([^/]*)/>")


def silent_multirun(xml: str) -> int:
    """Paragraphs with 2+ runs where at least one declares no weight."""
    hits = 0
    for body in PARA.findall(xml):
        runs = RUN.findall(body)
        if len(runs) < 2:
            continue
        for run in runs:
            m = re.search(r"<a:rPr\b[^>]*>", run)
            if m and ' b="' not in m.group(0):
                hits += 1
                break
    return hits


def deck_row(path: Path) -> tuple[int, bool]:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        bold_level = any(
            'b="1"' in z.read(n).decode("utf-8", "replace")
            for n in names
            if n.startswith(("ppt/slideLayouts/slideLayout", "ppt/slideMasters/"))
            and n.endswith(".xml")
        )
        hits = sum(
            silent_multirun(z.read(n).decode("utf-8", "replace"))
            for n in names
            if re.match(r"ppt/slides/slide\d+\.xml$", n)
        )
    return hits, bold_level


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--blind", action="store_true")
    args = ap.parse_args()
    src = ROOT / ("pptx" if args.blind else "dev/pptx")
    decks = sorted(src.glob("*.pptx"))
    reached = []
    for path in decks:
        try:
            hits, bold_level = deck_row(path)
        except Exception as exc:
            print("%-40s ERROR %s" % (path.stem[:40], exc))
            continue
        if hits and bold_level:
            reached.append((path.stem, hits))
    for stem, hits in sorted(reached, key=lambda r: -r[1]):
        print("%-52s %3d paragraphs" % (stem[:52], hits))
    print("\n%d of %d decks can be reached: %s"
          % (len(reached), len(decks),
             ",".join(s.split("__")[0] for s, _ in sorted(reached))))


if __name__ == "__main__":
    main()
