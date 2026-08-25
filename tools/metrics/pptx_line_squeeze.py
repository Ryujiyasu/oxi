# -*- coding: utf-8 -*-
"""Compare a drawn line's extent against the font's own advances -- CAUTIOUSLY.

★READ THIS BEFORE TRUSTING A NUMBER FROM HERE. The per-character `origin`
values pymupdf reports are RECONSTRUCTED, not read off the page, and on a page
that carries two subsets answering to the SAME name they can be reconstructed
from the wrong one. That is not hypothetical: d15 page 2 has both
`BCDFEE+Barlow Light` and `BCDGEE+Barlow Light`, and origin-derived extents
there disagree with the PDF's own `/Widths` array by ~1%. Three sessions'
worth of "PowerPoint squeezes lines to fit" came out of exactly that gap and
were WRONG -- `/Widths` for the disputed line sums to 275.40pt, the subset's
`hmtx` agrees to the unit, and Oxi computes 275.25pt. Nothing is compressed.

So:
  * `/Widths` (or `/W` for a Type0 font) is what POSITIONS the glyphs and is
    the authority. The embedded file's `hmtx` normally agrees with it.
  * origin-derived extents are a cross-check, never evidence on their own.
  * a name shared by two subsets makes every line using it unmeasurable here;
    those lines are dropped rather than guessed at.
  * a PDF usually omits space glyphs, so a space's origin is synthesised
    outright -- per-character deltas on spaces are meaningless.

What it still answers honestly: whether a deck's drawn lines depart from its
own font metrics anywhere, once the ambiguous names are removed. Median ratio
per (deck, font, size) came out 0.993..1.002 across six dev decks, i.e. no
systematic departure -- which is the useful negative result.

Usage:
    python tools/metrics/pptx_line_squeeze.py [d15 d20 ...]
"""
from __future__ import annotations

import glob
import io
import sys
from collections import defaultdict
from pathlib import Path

import pymupdf
from fontTools.ttLib import TTFont

REPO = Path(__file__).resolve().parents[2]
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def subset_metrics(doc, page):
    """font name as spans report it -> (cmap, hmtx, unitsPerEm).

    ★A page can carry TWO subsets that report the SAME name -- d15 page 2 has
    both `BCDFEE+Barlow Light` and `BCDGEE+Barlow Light`. A span only tells us
    the name, so such a name cannot be attributed to one subset and every line
    using it must be dropped; keeping the last one silently measures some lines
    against the wrong glyph table and manufactures outliers.
    """
    out = {}
    seen = set()
    for xref, _, _, base, _, _, _ in page.get_fonts(full=True):
        name = base.split("+")[-1]
        try:
            _, _, _, buf = doc.extract_font(xref)
        except Exception:
            continue
        if not buf:
            continue
        try:
            t = TTFont(io.BytesIO(buf))
        except Exception:
            continue
        if name in seen:
            out[name] = None          # ambiguous: refuse to guess
            continue
        seen.add(name)
        out[name] = (t.getBestCmap(), t["hmtx"], t["head"].unitsPerEm)
    return {k: v for k, v in out.items() if v is not None}


def line_ratio(chars, metrics):
    """drawn advance sum / design advance sum, or None if not measurable."""
    keep = list(chars)
    while keep and keep[-1]["c"] == " ":
        keep.pop()
    if len(keep) < 6:
        return None
    # The line TOTAL is measured between two real glyph origins, so it stands
    # even though a PDF usually omits space glyphs and the viewer synthesises
    # their origins. Per-character deltas do NOT stand, which is why this
    # reports a line ratio and never a per-glyph one.
    # One size and one font for the whole line, else the comparison is muddled.
    if len({round(c["size"], 2) for c in keep}) != 1:
        return None
    fonts = {c["font"] for c in keep}
    if len(fonts) != 1:
        return None
    face = metrics.get(next(iter(fonts)))
    if not face:
        return None
    cmap, hmtx, upem = face
    size = keep[0]["size"]
    design = 0.0
    for c in keep:
        g = cmap.get(ord(c["c"]))
        if g is None:
            return None
        design += hmtx[g][0] / upem * size
    if design <= 0:
        return None
    last = hmtx[cmap[ord(keep[-1]["c"])]][0] / upem * size
    drawn = keep[-1]["origin"][0] - keep[0]["origin"][0] + last
    return drawn / design, design, drawn, "".join(c["c"] for c in keep)


def main() -> None:
    decks = sys.argv[1:]
    pdfs = []
    for f in sorted(glob.glob(str(DEV / "pdf" / "*.pdf"))):
        deck = Path(f).name.split("__")[0]
        if not decks or deck in decks:
            pdfs.append((deck, f))
    if not pdfs:
        sys.exit("no reference PDFs matched")
    rows = []
    per_deck = defaultdict(list)
    for deck, f in pdfs:
        doc = pymupdf.open(f)
        for pno in range(len(doc)):
            page = doc[pno]
            metrics = subset_metrics(doc, page)
            if not metrics:
                continue
            for blk in page.get_text("rawdict")["blocks"]:
                if blk.get("type") != 0:
                    continue
                for ln in blk["lines"]:
                    # Rotated text advances along a different axis; comparing
                    # its x-extent to an advance sum measures the rotation.
                    if tuple(round(v, 3) for v in ln.get("dir", (1, 0))) != (1.0, 0.0):
                        continue
                    chars = [c for sp in ln["spans"] for c in sp["chars"]]
                    # rawdict puts font/size on the span; copy them down.
                    for sp in ln["spans"]:
                        for c in sp["chars"]:
                            c["font"], c["size"] = sp["font"], sp["size"]
                    got = line_ratio(chars, metrics)
                    if got:
                        ratio, design, drawn, text = got
                        rows.append((ratio, deck, pno + 1, design, drawn, text))
                        per_deck[deck].append(ratio)
    if not rows:
        sys.exit("no measurable lines")
    rows.sort()
    n = len(rows)
    print(f"measurable lines: {n} across {len(per_deck)} decks\n")
    print("MOST SQUEEZED (drawn narrower than the font's own design advances)")
    print(f"{'ratio':>8}{'design':>9}{'drawn':>9}  deck p#   line")
    for ratio, deck, pg, design, drawn, text in rows[:20]:
        print(f"{ratio:8.4f}{design:9.2f}{drawn:9.2f}  {deck} p{pg}  {text[:46]!r}")
    mid = rows[n // 2][0]
    print(f"\nmedian ratio {mid:.4f}  (the baseline: a line under no pressure)")
    squeezed = [r for r in rows if r[0] < mid - 0.004]
    print(f"lines squeezed more than 0.4% below that baseline: {len(squeezed)}")
    if squeezed:
        worst = squeezed[0]
        print(f"deepest squeeze {worst[0]:.4f} = {(1 - worst[0] / mid) * 100:.2f}% below baseline")
        print(f"  {worst[1]} p{worst[2]}  {worst[5][:60]!r}")


if __name__ == "__main__":
    main()
