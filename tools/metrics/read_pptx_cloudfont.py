# -*- coding: utf-8 -*-
"""Read the cloud-font probe back: which face did PowerPoint put on the page?

A PDF span carries the name of the face that drew it, so the answer needs no
metric at all -- if PowerPoint resolved the requested family the span says so,
and if it substituted the span says what it substituted.

Also prints the same tally for any deck re-exported by
`export_pptx_cloudfont.py --deck`, next to the tally of that deck's STORED truth
PDF, so a drift between the two is visible in one line.

Usage: python tools/metrics/read_pptx_cloudfont.py
"""
from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
CLOUD = REPO / "pipeline_data" / "pptx_probes" / "cloudfont"
PDF = CLOUD / "probe_cloudfont.pdf"
RE_EXPORT = CLOUD / "reexport"
TRUTH = REPO / "pipeline_data" / "pptx_benchmark" / "dev" / "pdf"


def spans(path: Path):
    doc = pymupdf.open(path)
    try:
        for page in doc:
            for block in page.get_text("rawdict")["blocks"]:
                for line in block.get("lines", []):
                    for span in line["spans"]:
                        yield span, "".join(c["c"] for c in span.get("chars", []))
    finally:
        doc.close()


def tally(path: Path) -> Counter:
    counts: Counter = Counter()
    for span, text in spans(path):
        counts[span["font"]] += len(text)
    return counts


def main() -> None:
    if PDF.is_file():
        print(f"{'arm':12s} {'requested':26s} {'PowerPoint drew':30s} verdict")
        for span, text in spans(PDF):
            match = re.match(r"([a-z]+):", text)
            if not match:
                continue
            arm = match.group(1)
            drew = span["font"]
            # The PDF name is the PostScript name (ArialMT, Calibri-Bold);
            # a resolved arm has the family in it, ignoring spaces and hyphens.
            print(f"{arm:12s} {'':26s} {drew:30s}")
    else:
        print(f"(no probe PDF at {PDF})")

    if RE_EXPORT.is_dir():
        print("\nre-exported corpus decks (fresh vs stored truth):")
        for fresh in sorted(RE_EXPORT.glob("*.pdf")):
            stored = sorted(TRUTH.glob(f"{fresh.stem}__*.pdf"))
            print(f"  {fresh.stem}  fresh  {tally(fresh).most_common(6)}")
            if stored:
                print(f"  {fresh.stem}  truth  {tally(stored[0]).most_common(6)}")


if __name__ == "__main__":
    main()
