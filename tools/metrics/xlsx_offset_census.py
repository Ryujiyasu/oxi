"""Which workbooks are a constant shift away from Excel, and which way?

Some losses are content — a fill missing, a rule not drawn — and some are the
same picture put down a pixel out. The second kind is cheap to chase when it
is systematic, and expensive to guess at when it is not, so it is worth
knowing which books are which before opening any of them.

Shifts are read from GLYPH BLOBS, not from correlating the ink profiles.
Correlation lies on Japanese body text: the characters sit on a regular pitch
and the profile matches itself a whole character away, so a book that is
perfectly aligned reports a confident shift of one character. Blob positions
cannot do that — they are compared in order, and the median only counts when
most of them agree on it.

    python tools/metrics/xlsx_offset_census.py            # the whole corpus
    python tools/metrics/xlsx_offset_census.py --worst 40 # the lowest 40 only
"""

from __future__ import annotations

import argparse
import json
import statistics
import sys
from pathlib import Path

import numpy as np
from PIL import Image

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
SHOTS = REPO / "pipeline_data" / "xlsx_diff_corpus"
BASELINE = REPO / "pipeline_data" / "xlsx_ssim_baseline.json"


def rows_of(ink: np.ndarray) -> list[tuple[int, int]]:
    lit = ink.sum(axis=1)
    out, start = [], None
    for at, held in enumerate(lit):
        if held > 2 and start is None:
            start = at
        elif held <= 2 and start is not None:
            if at - start >= 6:
                out.append((start, at))
            start = None
    return out


def blobs(profile: np.ndarray) -> list[int]:
    out, start = [], None
    for at, held in enumerate(profile):
        if held > 0 and start is None:
            start = at
        elif held == 0 and start is not None:
            out.append(start)
            start = None
    if start is not None:
        out.append(start)
    return out


def told(stem: str) -> dict | None:
    truth, ours = SHOTS / f"{stem}.excel.png", SHOTS / f"{stem}.oxi.png"
    if not (truth.exists() and ours.exists()):
        return None
    a = np.asarray(Image.open(truth).convert("L"))
    b = np.asarray(Image.open(ours).convert("L"))
    high, wide = min(a.shape[0], b.shape[0]), min(a.shape[1], b.shape[1])
    ia, ib = a[:high, :wide] < 140, b[:high, :wide] < 140
    across, down, lines = [], [], 0
    for y0, y1 in rows_of(ia):
        here, there = blobs(ia[y0:y1].sum(axis=0)), blobs(ib[y0:y1].sum(axis=0))
        if len(here) < 4 or len(there) < 4:
            continue
        count = min(len(here), len(there))
        apart = [there[at] - here[at] for at in range(count)]
        middle = statistics.median(apart)
        # Only a line where most blobs agree says anything.
        if sum(1 for held in apart if held == middle) / count < 0.7:
            continue
        across.append(int(middle))
        lines += 1
        # The same line's own top, against the nearest Excel line.
        rows_here = np.where(ia[max(0, y0 - 6):y1 + 6].any(axis=1))[0]
        rows_there = np.where(ib[max(0, y0 - 6):y1 + 6].any(axis=1))[0]
        if len(rows_here) and len(rows_there):
            down.append(int(rows_there.min()) - int(rows_here.min()))
    if not across:
        return None
    return {
        "lines": lines,
        "dx": int(statistics.median(across)),
        "dx_share": sum(1 for held in across if held == statistics.median(across)) / len(across),
        "dy": int(statistics.median(down)) if down else 0,
        "dy_share": (sum(1 for held in down if held == statistics.median(down)) / len(down))
        if down else 0.0,
    }


def main() -> int:
    parse = argparse.ArgumentParser()
    parse.add_argument("--worst", type=int, help="only the N lowest-scoring workbooks")
    parse.add_argument("--min-lines", type=int, default=3)
    args = parse.parse_args()

    scores = json.loads(BASELINE.read_text())["scores"]
    order = sorted(scores.items(), key=lambda held: held[1])
    if args.worst:
        order = order[: args.worst]

    shifted, square = [], 0
    for stem, score in order:
        found = told(stem)
        if not found or found["lines"] < args.min_lines:
            continue
        if found["dx"] == 0 and found["dy"] == 0:
            square += 1
            continue
        shifted.append((score, stem, found))

    print(f"  {len(shifted)} workbooks sit off Excel by a whole pixel or more;"
          f" {square} are square on both axes")
    print("  score    dx (agree)   dy (agree)   lines   workbook")
    for score, stem, found in sorted(shifted):
        print(f"  {score:.4f}  {found['dx']:+4d} ({found['dx_share']:.0%})"
              f"   {found['dy']:+4d} ({found['dy_share']:.0%})"
              f"   {found['lines']:>5}   {stem[:40]}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
