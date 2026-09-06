# -*- coding: utf-8 -*-
"""Balanced two-column bands in a Word PDF: for every band (a run of two-column
lines between full-width lines or page edges), the row tops of each column and
the band's bottom, so column-balance hypotheses can be scored on real sections
instead of on probe arms.

Usage: python tools/metrics/_colbalance_census.py <word.pdf> [--show] [--pages=a,b]

A row is a text line; its height is the distance to the next row top in the
same column (the last row's height is the band bottom minus its top). The
band bottom is the top of the first full-width line below the band (or the
column's last row top + the column's median pitch when the band ends the page).
For each band the census prints Word's split (rows in column 1 / column 2) and
which candidate rule predicts it:
  count   : ceil(n/2) rows in column 1
  ge      : smallest k with height(col1) >= height(col2)
  half    : smallest k with height(col1) >= total/2
"""
import math
import sys
from collections import Counter

import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
pdf = sys.argv[1]
show = "--show" in sys.argv
pages = None
for a in sys.argv[2:]:
    if a.startswith("--pages="):
        pages = [int(x) for x in a.split("=")[1].split(",")]


def bands_of(page):
    width = page.rect.width
    mid = width / 2
    lines = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if not t:
                continue
            x0, y0, x1, y1 = l["bbox"]
            lines.append((round(y0, 1), round(x0, 1), round(x1, 1), t))
    lines.sort()
    xs = Counter(x0 for _, x0, _, _ in lines)
    lefts = [x for x, c in xs.most_common(3) if c >= 3]
    if len(lefts) < 2:
        return []
    col_left = sorted(lefts)[:2]
    # the body pitch: the most common advance between successive column-1 lines
    c1y = sorted(y for y, x0, _, _ in lines if abs(x0 - col_left[0]) < 3)
    pitch = Counter(round(b - a, 1) for a, b in zip(c1y, c1y[1:]) if 8 < b - a < 40).most_common(1)
    pitch = pitch[0][0] if pitch else 20.0

    def col_of(x0):
        # a column row starts at the column's left edge or an indent of 1-3 characters;
        # anything else (table cells, centred lines) is not a column row
        for ci, left in enumerate(col_left):
            for k in range(0, 4):
                if abs(x0 - (left + k * pitch * 0.56)) < 2.5:
                    return ci
        return None

    rows = []
    for y0, x0, x1, t in lines:
        if abs(x0 - col_left[0]) < 3 and x1 > mid + 20:
            rows.append((y0, "full", t))
        else:
            ci = col_of(x0)
            rows.append((y0, "c%d" % (ci + 1) if ci is not None else "other", t))
    bands = []
    cur = {"c1": [], "c2": []}
    for y0, kind, t in rows:
        if kind in ("full", "other"):
            if cur["c1"] and cur["c2"]:
                bands.append((cur["c1"], cur["c2"], y0))
            cur = {"c1": [], "c2": []}
        else:
            cur[kind].append((y0, t))
    # a band still open at the page bottom is a continuing section, not a balance
    return bands


def rule_pick(h1rows, h2rows_all):
    """h1rows: heights of the rows in column 1 order... we only know the actual
    split; build the full row sequence = col1 rows then col2 rows."""
    return None


def main():
    doc = fitz.open(pdf)
    plist = pages or list(range(1, len(doc) + 1))
    tally = Counter()
    for pno in plist:
        page = doc[pno - 1]
        for c1, c2, bottom in bands_of(page):
            tops1 = [y for y, _ in c1]
            tops2 = [y for y, _ in c2]
            if len(tops1) + len(tops2) < 3:
                continue
            # a balanced band has both columns starting at the same top
            if abs(tops1[0] - tops2[0]) > 3.0:
                continue
            pitch = Counter(round(b - a, 1) for a, b in zip(tops1, tops1[1:])).most_common(1)
            pitch = pitch[0][0] if pitch else None
            if bottom is None:
                if pitch is None:
                    continue
                bottom = max(tops1[-1], tops2[-1]) + pitch
            h1 = [round(b - a, 1) for a, b in zip(tops1, tops1[1:])] + [round(bottom - tops1[-1], 1)]
            h2 = [round(b - a, 1) for a, b in zip(tops2, tops2[1:])] + [round(bottom - tops2[-1], 1)]
            rows = h1 + h2  # the section's rows in order
            n1 = len(h1)
            total = sum(rows)
            # candidate rules (k = rows in column 1)
            def height(k):
                return sum(rows[:k]), sum(rows[k:])
            preds = {}
            preds["count"] = math.ceil(len(rows) / 2)
            preds["ge"] = next((k for k in range(1, len(rows)) if height(k)[0] >= height(k)[1] - 0.5), len(rows))
            preds["half"] = next((k for k in range(1, len(rows)) if height(k)[0] >= total / 2 - 0.5), len(rows))
            for name, k in preds.items():
                tally[(name, k == n1)] += 1
            if show or any(k != n1 for k in preds.values()):
                print("p%d band top=%.1f bottom=%.1f  word=%d/%d  col1=%s col2=%s  count=%d ge=%d half=%d %s" % (
                    pno, tops1[0], bottom, n1, len(h2), h1, h2, preds["count"], preds["ge"], preds["half"],
                    "" if all(k == n1 for k in preds.values()) else "<-- " + ",".join(n for n, k in preds.items() if k != n1) + " miss"))
                if show:
                    print("      c1: %s" % " | ".join(t[:10] for _, t in c1))
                    print("      c2: %s" % " | ".join(t[:10] for _, t in c2))
    for name in ("count", "ge", "half"):
        print("%-6s hit=%d miss=%d" % (name, tally[(name, True)], tally[(name, False)]))


if __name__ == "__main__":
    main()
