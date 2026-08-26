# -*- coding: utf-8 -*-
"""How much of each deck's SSIM floor is PowerPoint's JPEG export, not Oxi?

PowerPoint's PDF export re-encodes bitmaps with DCT. Where a slide's backdrop
is a large FLAT image, the reference therefore carries compression grain that
the source does not have -- and SSIM, which compares local variance, punishes a
correct flat render for not reproducing it.

d28 is the proved case (2026-08-26). Its backdrop is `media/image15.png`, drawn
full-slide. Over the 756,460 pixels where Oxi paints one flat colour:

    original PNG, resized to render size   std = [0.37, 0.36, 0.10]   flat
    Oxi                                    std = [0.00, 0.00, 0.00]   correct
    PowerPoint reference                   std = [5.11, 3.85, 4.91]
    the JPEG stored inside that reference  std = [4.67, 3.87, 4.60]

The reference's grain IS the stored JPEG's grain. Oxi matches the source; the
reference does not. d28 scores 0.7793 -- the corpus's worst deck by 0.15 -- and
essentially all of that gap is this.

So the floor's worst decks must be read with this column beside them: the decks
with a full-slide JPEG on every page are exactly the decks at the bottom. That
does NOT mean their whole gap is artefact -- a backdrop with real detail hides
the grain -- but a deck at 100% here needs a per-slide look before anyone
treats its score as a defect to chase.

Usage:
    python tools/metrics/pptx_reference_jpeg_share.py [d28 d22 ...]
"""
from __future__ import annotations

import glob
import sys
from pathlib import Path

import pymupdf

REPO = Path(__file__).resolve().parents[2]
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# A bitmap covering this much of the page is a backdrop, not an illustration.
COVERAGE = 0.6
MAX_PAGES = 25


def share(pdf_path: Path) -> tuple[int, int]:
    """(pages whose backdrop is a JPEG, pages examined)"""
    doc = pymupdf.open(pdf_path)
    pages = min(len(doc), MAX_PAGES)
    hits = 0
    for pno in range(pages):
        page = doc[pno]
        area = page.rect.width * page.rect.height
        for info in page.get_images(full=True):
            try:
                img = doc.extract_image(info[0])
            except Exception:
                continue
            if img["ext"] != "jpeg":
                continue
            if any(r.width * r.height > COVERAGE * area
                   for r in page.get_image_rects(info[0])):
                hits += 1
                break
    return hits, len(doc)


def main() -> None:
    wanted = set(sys.argv[1:])
    rows = []
    for f in sorted(glob.glob(str(DEV / "pdf" / "*.pdf"))):
        deck = Path(f).name.split("__")[0]
        if wanted and deck not in wanted:
            continue
        hits, total = share(Path(f))
        rows.append((hits / max(min(total, MAX_PAGES), 1), deck, hits, total))
    if not rows:
        sys.exit("no reference PDFs matched")
    rows.sort(reverse=True)
    print(f"{'deck':<6}{'JPEG-backdrop pages':>22}{'share':>8}")
    for frac, deck, hits, total in rows:
        if hits:
            print(f"{deck:<6}{f'{hits}/{min(total, MAX_PAGES)}':>22}{frac * 100:>7.0f}%")
    clean = [r[1] for r in rows if r[2] == 0]
    print(f"\nno JPEG backdrop at all: {len(clean)}/{len(rows)}  {' '.join(clean)}")


if __name__ == "__main__":
    main()
