# -*- coding: utf-8 -*-
"""Try candidate column-split rules against every probe at once.

Changing the engine and re-rendering to test an idea costs eight minutes of
build; the arithmetic costs nothing. Every probe's rows are known — their
heights were measured, not assumed — so a candidate rule can be scored against
all of them here first, and only a rule that already fits is worth building.

    python tools/metrics/_col_rule_fit.py

Reports each candidate against the text-only families and the table family
separately, because the rule that ships fits the first and no rule yet fits
both.
"""
from __future__ import annotations

import sys
from statistics import median

# Arial line heights, read off this engine's own layout dump and confirmed
# against Word's paragraph positions: a line is its point size x 1.1499.
LINE = 11.499
TALL = {20: 11.499, 21: 12.073, 22: 12.649, 23: 13.224, 24: 13.799}
# A table row measured 12.0 on the probes, and the table's top border adds 0.5
# to the line that precedes it. Both from Information(6), not assumed.
TABLE_ROW = 12.0
TABLE_EDGE = 0.5


def positions() -> list:
    """(name, row heights, Word's left-column row count)."""
    out = []
    # One tall row swept through an otherwise even run.
    word_pos = {8: 4, 10: 5, 12: 6}  # left count while the tall row is at or before n/2
    for n, boundary in word_pos.items():
        for at in range(n):
            rows = [TALL[24] if i == at else LINE for i in range(n)]
            out.append((f"pos_{n}_at{at}", rows, n // 2 + (1 if at > boundary else 0)))
    return out


def sizes() -> list:
    out = []
    for n in (6, 8, 10, 12, 14):
        for hp in (20, 21, 22, 23, 24):
            rows = [LINE] * (n - 1) + [TALL[hp]]
            out.append((f"col_{n}_{hp}", rows, n // 2 + (0 if hp == 20 else 1)))
    return out


def tables() -> list:
    """8 text lines with a table of r rows dropped in at position p.

    Word's left counts are the measured ones, minus the one-column paragraph
    that follows the run and shares the left edge.
    """
    word = {
        (1, 0): 4, (1, 2): 4, (1, 4): 6, (1, 6): 5, (1, 8): 5,
        (2, 0): 5, (2, 2): 5, (2, 4): 5, (2, 6): 6, (2, 8): 6,
        (3, 0): 5, (3, 2): 5, (3, 4): 6, (3, 6): 6, (3, 8): 6,
        (5, 0): 6, (5, 2): 7, (5, 4): 7, (5, 6): 7, (5, 8): 7,
    }
    out = []
    for (r, at), want in sorted(word.items()):
        rows = []
        for i in range(8):
            if i == at:
                # The table's top border lands on the line above it.
                if rows:
                    rows[-1] += TABLE_EDGE
                rows.extend([TABLE_ROW] * r)
            rows.append(LINE)
        if at >= 8:
            rows[-1] += TABLE_EDGE
            rows.extend([TABLE_ROW] * r)
        out.append((f"tab_{r}_at{at}", rows, want))
    return out


def greedy(rows: list, limit: float) -> tuple[int, float]:
    left, split = 0.0, 0
    for i, h in enumerate(rows[:-1]):
        if left + h > limit + 0.001:
            break
        left += h
        split = i + 1
    if split == 0:
        return 1, rows[0]
    return split, left


def shipped(rows: list) -> int:
    """What the engine does now: half, fill, grow by a row until the rest fits."""
    total = sum(rows)
    step = max(median(rows), 0.01)
    limit = total * 0.5
    for _ in range(len(rows) + 1):
        split, left = greedy(rows, limit)
        if total - left <= limit + 0.001:
            return split
        limit += step
    return len(rows) - 1


def half_plus_row(rows: list) -> int:
    """Half plus one row, filled once."""
    total = sum(rows)
    return greedy(rows, total * 0.5 + max(median(rows), 0.01) - 0.001)[0]


def plain_half(rows: list) -> int:
    """Half, filled once, no growth."""
    return greedy(rows, sum(rows) * 0.5)[0]


def level(rows: list) -> int:
    """The split whose taller column is shortest; ties to the shorter left."""
    total, run, best = sum(rows), 0.0, (float("inf"), 1)
    for i, h in enumerate(rows[:-1]):
        run += h
        cost = max(run, total - run)
        if cost < best[0] - 0.001:
            best = (cost, i + 1)
    return best[1]


def grow_once(rows: list) -> int:
    """Half, and at most ONE row of growth."""
    total = sum(rows)
    step = max(median(rows), 0.01)
    split, left = greedy(rows, total * 0.5)
    if total - left <= total * 0.5 + 0.001:
        return split
    return greedy(rows, total * 0.5 + step)[0]


CANDIDATES = {
    "shipped (half, grow until the rest fits)": shipped,
    "half plus one row, filled once": half_plus_row,
    "half, filled once": plain_half,
    "level the two columns": level,
    "half, at most one row of growth": grow_once,
}


def main() -> int:
    families = {"position": positions(), "size": sizes(), "table": tables()}
    width = max(len(k) for k in CANDIDATES)
    print(f"{'rule':{width}}  " + "  ".join(f"{k:>10}" for k in families))
    for label, rule in CANDIDATES.items():
        counts = []
        misses = {}
        for fam, probes in families.items():
            ok = 0
            for name, rows, want in probes:
                got = rule(rows)
                if got == want:
                    ok += 1
                else:
                    misses.setdefault(fam, []).append((name, want, got))
            counts.append(f"{ok:3}/{len(probes):<3}")
        print(f"{label:{width}}  " + "  ".join(f"{c:>10}" for c in counts))
        for fam, rows in misses.items():
            if len(rows) <= 6:
                for name, want, got in rows:
                    print(f"    {fam} {name}: word {want} rule {got}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
