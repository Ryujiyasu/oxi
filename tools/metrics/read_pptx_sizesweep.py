# -*- coding: utf-8 -*-
"""The horizontal scale PowerPoint set each arm of sizesweep.pptx at.

One line per arm: the /Widths sum against the width the TJ array actually
draws, fitted as one scale through the origin (the same question
`pptx_line_condense.py` asks of the corpus, on a deck built to answer it).

The columns that matter:

    s              the scale
    size*s         the size the layout behaves as if it used
    round-size*s   how far that is from the nearest integer point

If the corpus reading is right, the last column is near zero for every arm and
`s` tracks `round(size)/size`. If a residual survives at the exact integers,
the integer story is only part of the answer and the residual is a second term
to name.
"""
import importlib.util
import os
import statistics
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location(
    "lc", REPO / "tools" / "metrics" / "pptx_line_condense.py")
lc = importlib.util.module_from_spec(spec)
spec.loader.exec_module(lc)

pdf = REPO / "pipeline_data" / "pptx_probes" / "sizesweep" / "sizesweep.pdf"
if not pdf.exists():
    sys.exit(f"no {pdf} -- run export_pptx_sizesweep.py first")

print(f"{'page':>5}{'font':>16}{'size':>9}{'s':>10}{'size*s':>10}"
      f"{'round-size*s':>14}{'resid':>8}")
rows = []
for pno, base, size, text, (s, scaled, plain), width in lc.runs_of(pdf, 8):
    face = base.split("+")[-1]
    rows.append((pno, face, size, s, scaled))
    print(f"{pno:5}{face[:15]:>16}{size:9.3f}{s:10.5f}{size*s:10.4f}"
          f"{round(size) - size*s:+14.4f}{scaled:8.3f}")

if rows:
    ints = [r for r in rows if abs(r[2] - round(r[2])) < 1e-6]
    if ints:
        print(f"\nexact integer sizes only ({len(ints)} arms): "
              f"s mean {statistics.mean(r[3] for r in ints):.5f}, "
              f"min {min(r[3] for r in ints):.5f}, "
              f"max {max(r[3] for r in ints):.5f}")
        print("  a residual here is NOT the integer-layout term.")
