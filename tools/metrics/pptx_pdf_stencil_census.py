# -*- coding: utf-8 -*-
"""How much of the blind corpus's truth is drawn as STENCILS, not as text?

When PowerPoint's PDF export cannot embed a face it was asked for, it does not
fail and it does not simply substitute: it writes the run **twice**. Once as a
text-showing operator under an `ExtGState` with `/ca 0` -- invisible, there so the
page stays searchable -- and once as ink, a 2x2 one-bit image of the run's colour
stretched over the word box and shaped by a high-resolution 1-bit `/SMask`.

That matters because the two halves need not agree with each other, and neither
need agree with what PowerPoint puts on SCREEN. On blind 31 slide 2 the stencil
ink is Calibri (the substitute) while the line breaks are the embedded part's,
leaving a 21.24pt hole between two runs -- and PowerPoint's own `Slide.Export`
of that slide breaks the same paragraph into FOUR lines in a rounded face, with
no hole. A deck whose truth is stencilled is a deck whose PDF may not be a
picture of PowerPoint.

    python tools/metrics/pptx_pdf_stencil_census.py           # every deck
    python tools/metrics/pptx_pdf_stencil_census.py 31 12     # named decks

Prints per deck: pages, pages carrying invisible text, invisible/total text ops,
stencil images. `pptx_slide_png.py` + `pptx_truth_png_vs_pdf.py` are what settle
an affected deck.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
PDF_DIR = REPO / "pipeline_data" / "pptx_benchmark" / "ssim_pptx" / "ppt_pdf"
TOK = re.compile(r"/GS\d+ gs|\bq\b|\bQ\b|\)\s*Tj|\]\s*TJ")


def page_stats(doc: pymupdf.Document, pno: int) -> tuple[int, int, int]:
    """(invisible text ops, total text ops, stencil images) on one page."""
    pg = doc[pno]
    ext = doc.xref_get_key(pg.xref, "Resources/ExtGState")
    alpha: dict[str, float] = {}
    if ext and ext[0] == "dict":
        for m in re.finditer(r"/(GS\d+) (\d+) 0 R", ext[1]):
            ca = re.search(r"/ca ([\d.]+)", doc.xref_object(int(m.group(2))))
            if ca:
                alpha[m.group(1)] = float(ca.group(1))
    ca, stack, inv, tot = 1.0, [], 0, 0
    for t in TOK.findall(pg.read_contents().decode("latin-1")):
        if t == "q":
            stack.append(ca)
        elif t == "Q":
            ca = stack.pop() if stack else 1.0
        elif t.endswith("gs"):
            ca = alpha.get(t.split()[0][1:], ca)
        else:
            tot += 1
            inv += ca == 0
    stencil = sum(1 for im in pg.get_images(full=True) if im[2] <= 2 and im[3] <= 2)
    return inv, tot, stencil


def main() -> None:
    want = [f"{int(a):02d}" for a in sys.argv[1:]]
    print(f"{'doc':>4} {'pages':>5} {'inkpg':>5} {'invis/text':>12} {'stencils':>8}")
    flagged = []
    for pdf in sorted(PDF_DIR.glob("[0-9][0-9].pdf")):
        key = pdf.stem
        if want and key not in want:
            continue
        doc = pymupdf.open(pdf)
        inv = tot = sten = pages = 0
        for i in range(len(doc)):
            a, b, c = page_stats(doc, i)
            inv, tot, sten = inv + a, tot + b, sten + c
            pages += a > 0
        share = f"{inv}/{tot}"
        print(f"{key:>4} {len(doc):>5} {pages:>5} {share:>12} {sten:>8}", flush=True)
        if inv:
            flagged.append((key, pages, len(doc), inv, tot))
        doc.close()
    if flagged:
        print("\ndecks whose truth PDF stencils text:")
        for key, pages, n, inv, tot in sorted(flagged, key=lambda r: -r[3] / max(r[4], 1)):
            print(f"  {key}: {pages}/{n} pages, {inv} of {tot} text ops invisible")


if __name__ == "__main__":
    main()
