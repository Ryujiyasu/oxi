# -*- coding: utf-8 -*-
"""Read the vertical-anchor overflow probe: where does PowerPoint put the block?

The box is 72.0pt tall with 7.2pt insets, so the text area is 57.6pt -- exactly
one 32pt line (pitch 1.2 x 32 = 38.4pt) fits with room to spare and two do not.
For each arm the reader prints every baseline PowerPoint drew and the offset of
the block's TOP from the text area's top, then compares it with the two
candidate models:

    clamped     off = max(0, inner_h - block_h) / d      (d = 1 for b, 2 for ctr)
    unclamped   off =        (inner_h - block_h)  / d

Usage: python tools/metrics/read_pptx_anchorb.py
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
PDF = REPO / "pipeline_data" / "pptx_probes" / "anchorb" / "probe_anchorb.pdf"

BOX_TOP = 144.0        # 2.00in
BOX_H = 72.0           # 1.00in
INS = 7.2              # 0.10in
SIZE = 32.0
PITCH = 1.2 * SIZE     # 38.4
INNER_H = BOX_H - 2 * INS
AREA_TOP = BOX_TOP + INS


def main() -> None:
    doc = pymupdf.open(PDF)
    print(f"box {BOX_TOP}..{BOX_TOP + BOX_H}  text area {AREA_TOP}..{AREA_TOP + INNER_H} "
          f"(inner {INNER_H})  line pitch {PITCH}")
    print(f"{'arm':8s} {'lines':5s} {'block_h':>8s} {'baselines':>34s} {'top_off':>8s} "
          f"{'clamped':>8s} {'unclamped':>10s}  verdict")
    for page in doc:
        spans = []
        for b in page.get_text("rawdict")["blocks"]:
            for line in b.get("lines", []):
                for s in line["spans"]:
                    t = "".join(c["c"] for c in s["chars"])
                    if re.fullmatch(r"(t|ctr|b)\dL\d", t):
                        spans.append((s["origin"][1], t))
        if not spans:
            continue
        spans.sort()
        arm = re.match(r"(t|ctr|b)", spans[0][1]).group(1)
        n = len(spans)
        block_h = n * PITCH
        # baseline offset inside a line box, from the single-line arm's own
        # geometry: for n=1 the block sits at a known place in every model.
        first = spans[0][0]
        d = 2.0 if arm == "ctr" else 1.0
        clamped = max(0.0, INNER_H - block_h) / d if arm != "t" else 0.0
        unclamped = (INNER_H - block_h) / d if arm != "t" else 0.0
        # base_off = first baseline measured from the block top, taken from the
        # t arm (which has no anchor offset at all).
        base_off = BASE_OFF.get(n)
        top_off = first - AREA_TOP - (base_off if base_off is not None else 0.0)
        if arm == "t":
            BASE_OFF[n] = first - AREA_TOP
            verdict = "(reference)"
            top_off = 0.0
        else:
            verdict = ("clamped" if abs(top_off - clamped) < 0.15 else
                       "UNCLAMPED" if abs(top_off - unclamped) < 0.15 else "neither")
        bl = " ".join(f"{y:.2f}" for y, _ in spans)
        print(f"{arm:8s} {n:5d} {block_h:8.2f} {bl:>34s} {top_off:8.2f} "
              f"{clamped:8.2f} {unclamped:10.2f}  {verdict}")


BASE_OFF: dict[int, float] = {}

if __name__ == "__main__":
    main()
