"""Which workbooks are missing a long straight run of ink, and where?

The merged-block border rule (SX66, 58 workbooks improved) was not found by
looking at pictures — it was found by counting ink per row. `glossary_05` read
"y=136: Excel 1873 px, Oxi 1128" and the columns Excel alone had lit turned
out to be one continuous 745-pixel run, which the layout dump placed on a row
boundary across one column: a merged block's bottom border.

This does that sweep over the whole corpus at once. For every workbook it
reports the longest run of pixels that Excel lights on a single row or column
and Oxi does not, and the longest the other way about. A long run is a rule,
a connector or a border — something structural — where a scatter of short
ones is glyph noise.

Reads the pictures the pixel diff already wrote; it renders nothing.

    python tools/metrics/xlsx_missing_rules.py
    python tools/metrics/xlsx_missing_rules.py --least 60
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import numpy as np
from PIL import Image

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
SHOTS = REPO / "pipeline_data" / "xlsx_diff_corpus"


def longest_run(mask: np.ndarray, least: int) -> tuple[int, int, int, str]:
    """The longest run of set pixels along a row, then along a column."""
    best = (0, 0, 0, "row")
    for axis, name in ((1, "row"), (0, "column")):
        lines = mask if axis == 1 else mask.T
        for index in range(lines.shape[0]):
            line = lines[index]
            if line.sum() < least:
                continue
            # Run-length over the lit stretches of this line.
            edges = np.flatnonzero(np.diff(np.concatenate(([0], line.view(np.int8), [0]))))
            starts, stops = edges[::2], edges[1::2]
            if not len(starts):
                continue
            runs = stops - starts
            at = int(runs.argmax())
            if runs[at] > best[0]:
                best = (int(runs[at]), index, int(starts[at]), name)
    return best


def main() -> int:
    parse = argparse.ArgumentParser()
    parse.add_argument("--least", type=int, default=80,
                       help="ignore runs shorter than this many pixels")
    parse.add_argument("--show", type=int, default=25)
    args = parse.parse_args()

    found = []
    for truth_png in sorted(SHOTS.glob("*.excel.png")):
        stem = truth_png.name[: -len(".excel.png")]
        ours_png = SHOTS / f"{stem}.oxi.png"
        if not ours_png.exists():
            continue
        a = np.asarray(Image.open(truth_png).convert("L"))
        b = np.asarray(Image.open(ours_png).convert("L"))
        high, wide = min(a.shape[0], b.shape[0]), min(a.shape[1], b.shape[1])
        ink_a, ink_b = a[:high, :wide] < 140, b[:high, :wide] < 140
        missing = longest_run(ink_a & ~ink_b, args.least)
        extra = longest_run(ink_b & ~ink_a, args.least)
        if missing[0] >= args.least or extra[0] >= args.least:
            found.append((stem, missing, extra))

    found.sort(key=lambda held: -max(held[1][0], held[2][0]))
    print(f"{len(found)} workbooks have a straight run of {args.least}px or more"
          f" that the other does not")
    print(f"{'workbook':<46}{'Excel alone':>26}{'Oxi alone':>26}")
    for stem, missing, extra in found[: args.show]:
        says = lambda run: (f"{run[0]:4d}px {run[3]} {run[1]} at {run[2]}" if run[0] else "  -")
        print(f"  {stem[:44]:<44}{says(missing):>26}{says(extra):>26}")
    out = SHOTS / "_missing_rules.json"
    out.write_text(json.dumps(
        [{"book": s, "excel_only": m, "oxi_only": e} for s, m, e in found], indent=1
    ), encoding="utf-8")
    print(f"written to {out}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
