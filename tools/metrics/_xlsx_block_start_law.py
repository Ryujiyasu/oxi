# -*- coding: utf-8 -*-
r"""Which of the face's own numbers carries the block's starting fraction?

`_xlsx_shape_pitch_size.py` reads twenty-four line tops an arm and solves the
pair (pitch, start) that Excel must have used: the pitch comes out within a
hundredth of the model's, and the start comes out on a FRACTION of a pixel —
so a block that begins on a whole one, which is what the renderer does, cannot
reproduce Excel's sequence. Fifteen of the arms rule a whole-pixel start out
outright.

This holds each candidate fraction against every arm's solved interval. A
candidate is only a law if it lands inside all of them.

    python tools\metrics\_xlsx_block_start_law.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import win32con
import win32gui
import win32ui

SOLVED = Path(r"C:\tmp\xlsx_shape_pitch_size\solved.json")
BIG = 2048
# The default `tIns` of a shape's body, which the probe's shapes do not state:
# 45720 EMU is 3.6pt, which is 4.8 pixels — and the renderer rounds it to 5
# before anything else is added to it.
INSET = 45720 / 9525.0


def ratios(face: str) -> tuple[float, float]:
    screen = win32gui.GetDC(0)
    dc = win32ui.CreateDCFromHandle(screen)
    font = win32ui.CreateFont({
        "name": face, "height": -BIG, "weight": win32con.FW_NORMAL,
        "charset": win32con.DEFAULT_CHARSET,
    })
    old = dc.SelectObject(font)
    told = dc.GetTextMetrics()
    dc.SelectObject(old)
    win32gui.ReleaseDC(0, screen)
    return told["tmHeight"] / BIG, told["tmAscent"] / BIG


def phase(value: float) -> float:
    """A fraction folded into the half-open pixel around zero."""
    held = value - round(value)
    return held


def main() -> int:
    held = json.loads(SOLVED.read_text(encoding="utf-8"))
    seen: dict[str, tuple[float, float]] = {}
    names: dict[str, int] = {}
    rows = []
    for arm in held:
        if arm["region"] is None:
            continue
        face, points = arm["face"], arm["points"]
        if face not in seen:
            seen[face] = ratios(face)
        tall, up = seen[face]
        em = points * 96.0 / 72.0
        natural = tall * em
        # The pitch the renderer models, and the half-leading it puts above
        # the first line.
        pitch = natural * 1.3
        lead = (pitch - natural) / 2.0
        ascent = up * em
        low, high = arm["region"][2], arm["region"][3]
        # Every candidate is a fraction of a pixel; what differs is which
        # numbers were allowed to keep theirs.
        candidates = {
            "0 (whole pixel)": 0.0,
            "inset": phase(INSET),
            "lead": phase(lead),
            "inset+lead": phase(INSET + lead),
            "inset+lead+asc": phase(INSET + lead + ascent),
            "lead+asc": phase(lead + ascent),
            "asc": phase(ascent),
            "inset+asc": phase(INSET + ascent),
            "inset+2lead": phase(INSET + 2 * lead),
            "natural": phase(natural),
            "inset+natural": phase(INSET + natural),
        }
        rows.append((face, points, low, high, candidates))
        for name in candidates:
            names.setdefault(name, 0)

    print(f"{'face':<14}{'pt':>6}{'solved f':>18}   "
          + "".join(f"{name:>17}" for name in names))
    for face, points, low, high, candidates in rows:
        marks = []
        for name in names:
            value = candidates[name]
            inside = low <= value < high or low <= value + 1.0 < high or \
                low <= value - 1.0 < high
            names[name] += inside
            marks.append(f"{value:+.3f}{'*' if inside else ' '}".rjust(17))
        print(f"{face:<14}{points:>6}{f'[{low:+.3f},{high:+.3f})':>18}   "
              + "".join(marks))
    print(f"\n{'':<20}{'held by':>18}   "
          + "".join(f"{names[name]:>17}" for name in names))
    print(f"{'':<20}{f'of {len(rows)} arms':>18}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
