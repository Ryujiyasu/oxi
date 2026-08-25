# -*- coding: utf-8 -*-
"""How far will PowerPoint SQUEEZE a line rather than move its last word down?

d15 slide 2 is the case that exposed this. Its right column sets

    that says "Download as PowerPoint template". You will

whose natural advance sum is 275.40pt in a 273.47pt box -- 1.93pt too wide --
and PowerPoint keeps it anyway, drawing it at 272.44pt. The squeeze is not
uniform: spaces give up 15% (2.400pt -> 2.040pt) while letters give up 2-3%.
Every OTHER line in the same shape is drawn +0.1..+0.3% WIDER than its design
sum, so this is not a measurement offset -- it is applied to the one line that
would not otherwise fit.

Oxi has no such allowance, so it breaks before "will" and re-flows the rest of
the paragraph. Before an allowance can be implemented its CAP has to be known,
which is what this measures: for every line PowerPoint drew, the ratio of the
drawn pen-advance sum to the design sum of the same characters.

Design advances come from the SUBSET PowerPoint embedded in its own PDF, so
they are the advances that font really has -- not a guess from a font of the
same name, which for these decks is often a different cut
([[pptx-embedded-part-slot-law]]).

★A line is only evidence about the cap if it was under pressure. A line with
room to spare is drawn at its natural width and says nothing, so the summary
reports the squeezed tail separately from the bulk.

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
    """font name as spans report it -> (cmap, hmtx, unitsPerEm)."""
    out = {}
    for xref, _, _, base, _, _, _ in page.get_fonts(full=True):
        try:
            _, _, _, buf = doc.extract_font(xref)
        except Exception:
            continue
        if not buf:
            continue
        try:
            t = TTFont(io.BytesIO(buf))
            out[base.split("+")[-1]] = (t.getBestCmap(), t["hmtx"], t["head"].unitsPerEm)
        except Exception:
            continue
    return out


def line_ratio(chars, metrics):
    """drawn advance sum / design advance sum, or None if not measurable."""
    keep = list(chars)
    while keep and keep[-1]["c"] == " ":
        keep.pop()
    if len(keep) < 6:
        return None
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
