# -*- coding: utf-8 -*-
"""Read the DRAW-GRID probe back out of PowerPoint's PDF.

Each arm is 40 copies of one glyph in a box far wider than the text, so the
drawn advance is the slope through their pen positions and the two rival models
are a single number apart:

    exact   em * sz
    grid    round(em * sz * 8) / 8      (the master unit, 576 to the inch)

The advance is read from the font file itself (fontTools), not from the engine,
so the probe does not inherit whatever the tables happen to say.

★THE ANSWER (2026-09-01): NEITHER. 18 of 24 arms land on neither model, and the
raw steps say why -- PowerPoint does not draw a repeated glyph with a constant
advance at all. Arial 12pt alternates 6.600 and 6.708pt along a single run of
40 identical letters; Arial 18pt is 9.984 throughout with one 10.0922. There is
no wrapping, no alignment, no box to overflow and one run, so whatever varies
the step is NOT the line squeeze answering a demand -- which is how the corpus's
long-line residual has been read until now.

Usage: python tools/metrics/read_pptx_drawgrid.py
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import pymupdf
from fontTools.ttLib import TTFont

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_probes" / "drawgrid"
FILES = {"Arial": "arial.ttf", "Georgia": "georgia.ttf", "Verdana": "verdana.ttf"}


def advance_em(face: str, ch: str) -> float:
    """The face's own advance for `ch`, in EM units, from the font file."""
    f = TTFont(Path(os.environ["WINDIR"]) / "Fonts" / FILES[face],
               lazy=True, checkChecksums=0)
    upm = f["head"].unitsPerEm
    cmap = f.getBestCmap()
    name = cmap[ord(ch)]
    return f["hmtx"][name][0] / upm


def page_runs(page):
    """Characters grouped by baseline, in reading order, with pen positions."""
    rows: dict[float, list] = {}
    for block in page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                for ch in span.get("chars", []):
                    rows.setdefault(round(span["origin"][1], 2), []).append(
                        {"c": ch["c"], "x": ch["origin"][0], "size": span["size"]})
    return [sorted(v, key=lambda c: c["x"]) for _, v in sorted(rows.items())]


def main() -> None:
    arms = json.loads((OUT / "arms.json").read_text(encoding="utf-8"))
    doc = pymupdf.open(OUT / "probe_drawgrid.pdf")
    print(f"{'face':10}{'sz':>7}{'em':>9}{'exact':>9}{'grid':>9}"
          f"{'drawn':>9}{'|d-exact|':>11}{'|d-grid|':>10}  verdict")
    agree = {"grid": 0, "exact": 0, "neither": 0}
    for arm in arms:
        page = doc[arm["slide"] - 1]
        runs = page_runs(page)
        rep = next((r for r in runs if len(r) >= 30
                    and all(c["c"] == arm["repeat"][0] for c in r)), None)
        if not rep:
            print(f"{arm['typeface']:10}{arm['sz_pt']:7.2f}   no repeated run on the page")
            continue
        sz = arm["sz_pt"]
        em = advance_em(arm["typeface"], arm["repeat"][0])
        exact = em * sz
        grid = round(em * sz * 8.0) / 8.0
        # The slope through every glyph, not just the ends: a single reading
        # error then cannot decide the arm.
        n = len(rep)
        xs = [c["x"] - rep[0]["x"] for c in rep]
        drawn = sum(x * i for i, x in enumerate(xs)) / sum(i * i for i in range(n))
        de, dg = abs(drawn - exact), abs(drawn - grid)
        # A verdict needs the two models to be far enough apart to have one.
        if abs(exact - grid) < 0.004:
            verdict = "(models agree)"
        elif dg < de and dg < 0.004:
            verdict = "GRID"
            agree["grid"] += 1
        elif de < dg and de < 0.004:
            verdict = "exact"
            agree["exact"] += 1
        else:
            verdict = "neither"
            agree["neither"] += 1
        print(f"{arm['typeface']:10}{sz:7.2f}{em:9.5f}{exact:9.4f}{grid:9.4f}"
              f"{drawn:9.4f}{de:11.4f}{dg:10.4f}  {verdict}")
    print(f"\n{agree['grid']} arms on the grid, {agree['exact']} on the exact "
          f"advance, {agree['neither']} on neither")

    # The kerned pair, which the single-glyph arms cannot speak about.
    print("\nthe alternating pair, first six steps per arm:")
    for arm in arms[:6]:
        page = doc[arm["slide"] - 1]
        runs = page_runs(page)
        alt = next((r for r in runs if len(r) >= 30
                    and r[0]["c"] == arm["alternate"][0]
                    and r[1]["c"] == arm["alternate"][1]), None)
        if not alt:
            continue
        steps = [alt[i + 1]["x"] - alt[i]["x"] for i in range(6)]
        pair = [advance_em(arm["typeface"], c) * arm["sz_pt"]
                for c in arm["alternate"][:2]]
        gridp = [round(p * 8.0) / 8.0 for p in pair]
        print(f"  {arm['typeface']:10}{arm['sz_pt']:6.2f}  drawn "
              + " ".join(f"{s:7.4f}" for s in steps)
              + f"   exact {pair[0]:.4f}/{pair[1]:.4f}"
              + f"   grid {gridp[0]:.4f}/{gridp[1]:.4f}")
    doc.close()


if __name__ == "__main__":
    main()
