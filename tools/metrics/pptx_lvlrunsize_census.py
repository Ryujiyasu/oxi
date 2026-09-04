# -*- coding: utf-8 -*-
"""Does a run that names no SIZE take the level's, or the paragraph's?

`paragraph_font_size` returns the largest size any run declares, and every run
that declares none is then measured at THAT. PowerPoint appears to give a silent
run the LEVEL's size instead, which is a different number whenever one sibling
run carries an `sz` and the level says something else.

d19 slide 4 is the specimen. Its title is three runs -- `1`, `.` (`sz="2400"`),
` Transition headline` -- and its truth PDF draws them at 45.98 / 24.00 / 45.98.
The engine measures the whole line at 24 and comes out 88.8% narrow.

This finds every paragraph of that shape and asks the deck's truth PDF what
sizes it actually drew there, so the rule is answered by PowerPoint rather than
inferred from one slide.

    python tools/metrics/pptx_lvlrunsize_census.py [--blind] [--decks d19,d22]
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[2] / "pipeline_data" / "pptx_benchmark"
PARA = re.compile(r"<a:p>(.*?)</a:p>", re.S)
RUN = re.compile(r"<a:r>(.*?)</a:r>", re.S)


def mixed_paragraphs(xml: str) -> list[tuple[list[str], list[float | None]]]:
    """Paragraphs where one run declares a size and another does not."""
    out = []
    for body in PARA.findall(xml):
        texts, sizes = [], []
        for run in RUN.findall(body):
            rpr = re.search(r"<a:rPr\b[^>]*>", run)
            sz = re.search(r'sz="(\d+)"', rpr.group(0)) if rpr else None
            text = re.search(r"<a:t>(.*?)</a:t>", run, re.S)
            texts.append(re.sub(r"\s+", " ", text.group(1)) if text else "")
            sizes.append(int(sz.group(1)) / 100.0 if sz else None)
        if len(sizes) >= 2 and any(s is None for s in sizes) and any(s for s in sizes):
            out.append((texts, sizes))
    return out


def page_sizes(page, needle: str) -> list[tuple[str, float]]:
    """(text, size) for the spans on this page that carry the needle."""
    got = []
    for blk in page.get_text("dict")["blocks"]:
        for line in blk.get("lines", []):
            for span in line["spans"]:
                if needle and needle[:12] in span["text"]:
                    got.append((span["text"][:26], round(span["size"], 2)))
    return got


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--blind", action="store_true")
    ap.add_argument("--decks", default="")
    args = ap.parse_args()
    src = ROOT / ("pptx" if args.blind else "dev/pptx")
    pdfs = ROOT / ("ssim_pptx/ppt_pdf" if args.blind else "dev/pdf")
    want = {d.strip() for d in args.decks.split(",") if d.strip()}

    decks = 0
    paragraphs = 0
    for path in sorted(src.glob("*.pptx")):
        if want and path.stem.split("__")[0] not in want:
            continue
        pdf_hits = sorted(pdfs.glob(path.stem + "*.pdf"))
        doc = pymupdf.open(pdf_hits[0]) if pdf_hits else None
        printed = False
        try:
            with zipfile.ZipFile(path) as z:
                slides = sorted(
                    (int(re.search(r"slide(\d+)\.xml$", n).group(1)), n)
                    for n in z.namelist()
                    if re.match(r"ppt/slides/slide\d+\.xml$", n))
                for num, name in slides:
                    xml = z.read(name).decode("utf-8", "replace")
                    for texts, sizes in mixed_paragraphs(xml):
                        silent = next((t for t, s in zip(texts, sizes)
                                       if s is None and len(t.strip()) > 4), "")
                        stated = next(s for s in sizes if s)
                        drawn = page_sizes(doc[num - 1], silent) if doc and num <= doc.page_count else []
                        if not drawn:
                            continue
                        sizes_drawn = {d for _t, d in drawn}
                        if sizes_drawn == {stated}:
                            continue  # the silent run took the sibling's size
                        paragraphs += 1
                        if not printed:
                            print(path.stem[:48])
                            printed = True
                        print("   s%-3d silent %-26r drawn at %s, sibling declares %.2f"
                              % (num, silent[:26], sorted(sizes_drawn), stated))
        finally:
            if doc:
                doc.close()
        decks += printed
    print("\n%d paragraphs over %d decks where the silent run is NOT drawn at "
          "its sibling's size" % (paragraphs, decks))


if __name__ == "__main__":
    main()
