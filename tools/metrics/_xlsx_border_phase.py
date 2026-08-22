# -*- coding: utf-8 -*-
"""What anchors the phase of a broken rule?

`_xlsx_border_pattern.py` showed a hair rule is inked exactly where `(x + y)`
is even, everywhere on the sheet — which is a halftone pinned to the picture,
not a run started at the cell. This fits the same question for every broken
style: it reads each rule off Excel's picture, throws away the pixels where
another rule crosses, and tests the candidate anchors against every rule at
once.

    python tools\\metrics\\_xlsx_border_phase.py            (reuses the pictures)
"""
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image

REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SCRATCH = Path(r"C:\tmp\xlsx_border")
STYLES = ["hair", "dotted", "dashed", "dashDot", "dashDotDot", "mediumDashed"]


def geometry(path):
    import os

    environment = dict(os.environ, OXI_XLSX_DUMP_COLUMNS="1", OXI_XLSX_DUMP_ROWS="1")
    done = subprocess.run([str(RENDERER), str(path), str(SCRATCH / "phase.oxi.png"), "96"],
                          capture_output=True, timeout=300, env=environment)
    columns, rows = {}, {}
    for line in done.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "column":
            columns[int(parts[1])] = int(float(parts[3]))
        elif len(parts) == 4 and parts[0] == "row":
            rows[int(parts[1])] = int(float(parts[3]))

    def edges(sizes):
        out, at = [0], 0
        for index in sorted(sizes):
            at += sizes[index]
            out.append(at)
        return out

    return edges(columns), edges(rows)


def samples(truth, across, down):
    """(x, y, inked) for the pixels of the rules each way, minus the crossings."""
    height, width = truth.shape
    flat, upright = [], []
    for y in down[:-1]:
        if y >= height:
            continue
        for x in range(across[0], min(across[-1], width)):
            if any(abs(x - edge) <= 1 for edge in across):
                continue
            flat.append((x, y, bool(truth[y, x] < 128)))
    for x in across[:-1]:
        if x >= width:
            continue
        for y in range(down[0], min(down[-1], height)):
            if any(abs(y - edge) <= 1 for edge in down):
                continue
            upright.append((x, y, bool(truth[y, x] < 128)))
    return flat, upright


def fit(points, key, period):
    """Is `inked` a function of `key(x, y) % period`? Return the inked set."""
    inked, clear = set(), set()
    for x, y, on in points:
        (inked if on else clear).add(key(x, y) % period)
    return None if inked & clear else inked


ANCHORS = [
    ("x + y", lambda x, y: x + y),
    ("x - y", lambda x, y: x - y),
    ("x + 2y", lambda x, y: x + 2 * y),
    ("2x + y", lambda x, y: 2 * x + y),
    ("x", lambda x, y: x),
    ("y", lambda x, y: y),
]
PERIODS = (2, 3, 4, 6, 8, 9, 12, 16, 18, 20, 24, 32, 36)


def main():
    for style in STYLES:
        path = SCRATCH / f"border_{style}.xlsx"
        picture = path.with_suffix(".excel.png")
        if not picture.exists():
            print(f"{style}: (no picture — run _xlsx_border_pattern.py first)")
            continue
        truth = np.asarray(Image.open(picture).convert("L"))
        across, down = geometry(path)
        flat, upright = samples(truth, across, down)
        for way, points in (("across", flat), ("down", upright), ("both", flat + upright)):
            found = []
            for period in PERIODS:
                for name, key in ANCHORS:
                    inked = fit(points, key, period)
                    if inked is not None:
                        found.append(f"({name}) mod {period} in {sorted(inked)}")
                if found:
                    break
            share = sum(1 for _, _, on in points if on) / max(1, len(points))
            print(f"{style:<14}{way:<8}{len(points):>5} pixels, {share:5.1%} inked  "
                  f"{found[0] if found else 'no single anchor fits'}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
