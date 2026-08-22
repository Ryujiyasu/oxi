# -*- coding: utf-8 -*-
"""Read the line-end repro from PowerPoint's PDF vectors, exactly.

The repro sweeps line width against the (w, len) size tokens for five end
types. Reading the PDF's filled PATHS rather than pixels makes each head's
extent exact, which is what the law needs: the whole question is whether the
size is a plain multiple of the line width or a multiple of a FLOORED width,
and the two differ by fractions of a point on thin lines.

Reported per connector: the head's extent across and along the line, both
divided by the line width and by the floored width, so the fit is readable.

Usage: python tools/metrics/read_pptx_lineend_probe.py [--floor 2.0]
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PROBE = Path(r"pipeline_data\pptx_probe\probe_lineend.pptx")
FACTOR = {"sm": 2.0, "med": 3.0, "lg": 5.0}


def path_points(d: dict) -> list[tuple[float, float]]:
    out = []
    for item in d["items"]:
        for p in item[1:]:
            if isinstance(p, pymupdf.Point):
                out.append((p.x, p.y))
            elif isinstance(p, pymupdf.Rect):
                out += [(p.x0, p.y0), (p.x1, p.y1)]
            elif isinstance(p, pymupdf.Quad):
                out += [(q.x, q.y) for q in (p.ul, p.ur, p.ll, p.lr)]
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--floor", type=float, default=2.0)
    args = ap.parse_args()

    manifest = json.loads(PROBE.with_suffix(".json").read_text(encoding="utf-8"))
    pdf = pymupdf.open(PROBE.with_suffix(".pdf"))
    by_slide: dict[int, list] = {}
    scale = None
    for row in manifest:
        s = row["slide"]
        if s not in by_slide:
            page = pdf[s - 1]
            scale = page.rect.width / (13.333 * 72)
            by_slide[s] = [d for d in page.get_drawings()
                           if d["type"] in ("f", "fs")]
        drawings = by_slide[s]
        x0 = row["x0_pt"] * scale
        y = row["y_pt"] * scale
        lw = row["line_pt"]
        # The head that owns this endpoint. A pointed head starts exactly at
        # x0; an OVAL is centred on it and so starts up to half a head before,
        # which is why a "path begins at x0" test finds every kind but that one.
        reach = FACTOR["lg"] * max(lw, args.floor) * scale
        best = None
        for d in drawings:
            r = d["rect"]
            if not (x0 - reach <= r.x0 <= x0 + 1.0):
                continue
            if not (r.y0 - 2 <= y <= r.y1 + 2):
                continue
            if r.width > row["x1_pt"] * scale:
                continue
            if best is None or r.width < best["rect"].width:
                best = d
        if best is None:
            print(f'{row["type"]:9s} {row["w_tok"]:4s} {row["len_tok"]:4s} '
                  f'{lw:5.2f}  (no path at the endpoint)')
            continue
        pts = path_points(best)
        # the head is everything within one head-length of the tip; the stem
        # runs the other way, so cut at the head's own extent
        head_len = FACTOR[row["len_tok"]] * max(lw, args.floor) * scale
        near = [p for p in pts if p[0] <= x0 + head_len + 0.51 and p[0] >= x0 - reach]
        across = (max(p[1] for p in near) - min(p[1] for p in near)) / scale
        along = (max(p[0] for p in near) - min(p[0] for p in near)) / scale
        past = (x0 - min(p[0] for p in near)) / scale
        fw, fl = FACTOR[row["w_tok"]], FACTOR[row["len_tok"]]
        print(f'{row["type"]:9s} {row["w_tok"]:4s} {row["len_tok"]:4s} {lw:5.2f} | '
              f'across {across:7.3f} = {across / lw:6.3f}xlw = '
              f'{across / (fw * max(lw, args.floor)):5.3f}x pred | '
              f'along {along:7.3f} = {along / (fl * max(lw, args.floor)):5.3f}x pred | '
              f'past {past:6.3f}')


if __name__ == "__main__":
    main()
