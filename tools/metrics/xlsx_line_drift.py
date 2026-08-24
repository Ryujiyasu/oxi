"""Does a line of a workbook's text drift as it runs?

`xlsx_offset_census.py` asks whether a whole book sits a pixel out. This asks a
different question of the same pictures: within ONE line of text, is the ink
further out of step at the end than at the start? That is what a wrong
per-character advance looks like — the line starts where Excel starts it and
loses a pixel every few characters — and it is invisible to a whole-book offset,
which sees only the median.

Each text band is split in half by its own ink. The best whole-pixel shift is
found for each half separately, searching only ±4 pixels so the character pitch
(13 to 16 pixels here) cannot alias one character onto the next. The drift is
the right half's shift less the left's.

    python tools/metrics/xlsx_line_drift.py               # the whole corpus
    python tools/metrics/xlsx_line_drift.py --worst 30    # the lowest 30 only
    python tools/metrics/xlsx_line_drift.py <stem>        # one workbook, per band
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
REACH = 4


def bands(ink: np.ndarray) -> list[tuple[int, int]]:
    """Runs of rows that carry text, kept apart by their blank rows."""
    lit = ink.sum(axis=1)
    out, start = [], None
    for at, held in enumerate(lit):
        if held > 8 and start is None:
            start = at
        elif held <= 8 and start is not None:
            if 6 <= at - start <= 40:
                out.append((start, at))
            start = None
    return out


def best_shift(theirs: np.ndarray, ours: np.ndarray) -> int | None:
    """The whole-pixel shift that puts the most of our ink on theirs."""
    if theirs.sum() < 40 or ours.sum() < 40:
        return None
    scores = []
    for shift in range(-REACH, REACH + 1):
        moved = np.roll(ours, shift)
        if shift > 0:
            moved[:shift] = False
        elif shift < 0:
            moved[shift:] = False
        scores.append(((theirs & moved).sum(), shift))
    best = max(scores)
    # A tie says nothing: the band has to prefer one shift to the others.
    if sum(1 for score, _ in scores if score == best[0]) > 1:
        return None
    return best[1]


def drift_of(stem: str, show: bool = False) -> tuple[float, int] | None:
    truth_png = SHOTS / f"{stem}.excel.png"
    ours_png = SHOTS / f"{stem}.oxi.png"
    if not (truth_png.exists() and ours_png.exists()):
        return None
    theirs = np.asarray(Image.open(truth_png).convert("L")) < 140
    ours = np.asarray(Image.open(ours_png).convert("L")) < 140
    if theirs.shape != ours.shape:
        return None
    drifts = []
    for top, bottom in bands(theirs):
        line = theirs[top:bottom]
        mine = ours[top:bottom]
        lit = np.where(line.any(axis=0))[0]
        if len(lit) < 200:
            continue
        middle = (int(lit.min()) + int(lit.max())) // 2
        left = best_shift(line[:, :middle].any(axis=0), mine[:, :middle].any(axis=0))
        right = best_shift(line[:, middle:].any(axis=0), mine[:, middle:].any(axis=0))
        if left is None or right is None:
            continue
        drifts.append(right - left)
        if show:
            print(f"    rows {top:>5}-{bottom:<5} x {int(lit.min()):>4}-{int(lit.max()):<4}"
                  f"  left {left:+d}  right {right:+d}  drift {right - left:+d}")
    if not drifts:
        return None
    return statistics.mean(drifts), len(drifts)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("stem", nargs="?", help="one workbook, band by band")
    parser.add_argument("--worst", type=int, help="only the N lowest-scoring workbooks")
    args = parser.parse_args()
    scores: dict[str, float] = json.loads(BASELINE.read_text(encoding="utf-8"))["scores"]

    if args.stem:
        print(f"  {args.stem}  SSIM {scores.get(args.stem, float('nan')):.4f}")
        got = drift_of(args.stem, show=True)
        print("  nothing to read" if got is None
              else f"  mean drift {got[0]:+.2f}px over {got[1]} bands")
        return 0

    stems = sorted(scores, key=lambda stem: scores[stem])
    if args.worst:
        stems = stems[: args.worst]
    told = []
    for stem in stems:
        got = drift_of(stem)
        if got is not None and got[1] >= 3:
            told.append((got[0], got[1], stem))
    told.sort(key=lambda row: -abs(row[0]))
    print(f"  {len(told)} workbook(s) read, worst drift first")
    print("   drift  bands  SSIM    workbook")
    for drift, count, stem in told[:40]:
        print(f"  {drift:+6.2f} {count:>6}  {scores[stem]:.4f}  {stem}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
