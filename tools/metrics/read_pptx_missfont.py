# -*- coding: utf-8 -*-
"""Read the missing-face probe back: which face did PowerPoint draw per arm?

The arm's label is drawn in the same run as its sample text, so the PDF span
that carries the label carries the answer. A miss that comes back as the THEME
face and a miss that comes back as Calibri are different rules, and this deck
is themed Georgia so that they cannot both fit.

Usage: python tools/metrics/read_pptx_missfont.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
PDF = REPO / "pipeline_data" / "pptx_probes" / "missfont" / "missfont.pdf"


def main() -> None:
    if not PDF.exists():
        sys.exit(f"{PDF} is not there -- run export_pptx_missfont.py first")
    doc = pymupdf.open(PDF)
    print(f"{PDF.name}: {doc.page_count} pages\n")
    for pno in range(doc.page_count):
        page = doc[pno]
        spans = [
            s
            for b in page.get_text("dict")["blocks"]
            for line in b.get("lines", [])
            for s in line["spans"]
            if s["text"].strip()
        ]
        for s in spans:
            face = s["font"].split("+")[-1]
            print(f"  p{pno + 1}  {face:<28} {s['size']:6.2f}  {s['text'][:44]!r}")
        if not spans:
            print(f"  p{pno + 1}  (no text spans -- the page drew nothing)")
    print("\nembedded font objects, per page:")
    for pno in range(doc.page_count):
        faces = sorted({f[3].split("+")[-1] for f in doc[pno].get_fonts()})
        print(f"  p{pno + 1}  {faces}")


if __name__ == "__main__":
    main()
