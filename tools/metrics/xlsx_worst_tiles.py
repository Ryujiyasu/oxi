"""Where in a workbook does the score actually go?

The long-run sweep (`xlsx_missing_rules.py`) finds structure — a rule, a
border, a connector. It says nothing about a region that is merely *wrong*:
text a shade off, a box with one line too many, a fill that stops early. For
that, score the picture in tiles and list the worst.

That is how `002`'s floor was traced: every one of its worst tiles sat in one
column band, which turned out to be the yellow notes down the right-hand
side, where Excel stops after three lines and Oxi drew a fourth.

Reads the pictures the pixel diff already wrote; it renders nothing.

    python tools/metrics/xlsx_worst_tiles.py b6a3a84180c9_002
    python tools/metrics/xlsx_worst_tiles.py --worst 8
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
SHOTS = REPO / "pipeline_data" / "xlsx_diff_corpus"
BASELINE = REPO / "pipeline_data" / "xlsx_ssim_baseline.json"


def tiles_of(stem: str, side: int, show: int) -> None:
    truth_png = SHOTS / f"{stem}.excel.png"
    ours_png = SHOTS / f"{stem}.oxi.png"
    if not (truth_png.exists() and ours_png.exists()):
        print(f"  {stem}: no pictures")
        return
    a = np.asarray(Image.open(truth_png).convert("L"))
    b = np.asarray(Image.open(ours_png).convert("L"))
    high, wide = min(a.shape[0], b.shape[0]), min(a.shape[1], b.shape[1])
    a, b = a[:high, :wide], b[:high, :wide]
    held = []
    for y in range(0, high - side, side):
        for x in range(0, wide - side, side):
            one, two = a[y:y + side, x:x + side], b[y:y + side, x:x + side]
            if one.std() < 1 and two.std() < 1:
                continue
            held.append((structural_similarity(one, two), y, x))
    if not held:
        print(f"  {stem}: nothing with content")
        return
    held.sort()
    # Which column band and which row band the loss sits in.
    columns: dict[int, list[float]] = {}
    rows: dict[int, list[float]] = {}
    for score, y, x in held:
        columns.setdefault(x, []).append(score)
        rows.setdefault(y, []).append(score)
    worst_column = min(columns.items(), key=lambda held: np.mean(held[1]))
    worst_row = min(rows.items(), key=lambda held: np.mean(held[1]))
    print(f"  {stem}  {wide}x{high}, {len(held)} tiles of {side}px")
    print(f"    the whole picture   {np.mean([s for s, _, _ in held]):.4f}")
    print(f"    worst column band   x {worst_column[0]:5d}  {np.mean(worst_column[1]):.4f}"
          f"  ({len(worst_column[1])} tiles)")
    print(f"    worst row band      y {worst_row[0]:5d}  {np.mean(worst_row[1]):.4f}"
          f"  ({len(worst_row[1])} tiles)")
    for score, y, x in held[:show]:
        print(f"      tile y {y:5d} x {x:5d}   {score:.4f}")


def main() -> int:
    parse = argparse.ArgumentParser()
    parse.add_argument("book", nargs="?", help="a workbook stem; omit to take the worst")
    parse.add_argument("--worst", type=int, default=3,
                       help="how many of the lowest-scoring workbooks to look at")
    parse.add_argument("--side", type=int, default=128)
    parse.add_argument("--show", type=int, default=8)
    args = parse.parse_args()

    if args.book:
        tiles_of(args.book, args.side, args.show)
        return 0
    scores = json.loads(BASELINE.read_text())["scores"]
    for stem, _ in sorted(scores.items(), key=lambda held: held[1])[: args.worst]:
        tiles_of(stem, args.side, args.show)
        print()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
