# -*- coding: utf-8 -*-
"""Compare the TEXT PowerPoint draws with the text Oxi draws, line by line.

The pixel metrics say a slide is wrong; this says WHERE. PowerPoint's exported
PDF carries every span with its position and size, and oxi-pptx-renderer's
`--dump-layout` carries the paragraphs it laid out, so the two can be matched
on their text and their y compared directly -- no ink thresholds, no blob
merging, none of the guesswork that reading a downscaled 3-panel invites.

Usage:
    python tools/metrics/pptx_text_diff.py d28 13
    python tools/metrics/pptx_text_diff.py d28 13 --tag sf --all
"""
from __future__ import annotations

import argparse
import glob
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = Path(__file__).resolve().parents[2]
DEV = REPO_ROOT / "pipeline_data" / "pptx_benchmark" / "dev"
EXE = REPO_ROOT / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"


def deck_path(prefix: str) -> Path:
    hits = sorted(glob.glob(str(DEV / "pptx" / f"{prefix}*.pptx")))
    if not hits:
        sys.exit(f"no deck matching {prefix}")
    return Path(hits[0])


def powerpoint_lines(pdf_path: Path, page: int) -> list[tuple[float, float, float, str]]:
    """(y_top, y_baseline-ish, size, text) per line, in PowerPoint's PDF."""
    doc = pymupdf.open(pdf_path)
    out = []
    for block in doc[page - 1].get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            text = "".join(s["text"] for s in line["spans"]).strip()
            if not text:
                continue
            spans = line["spans"]
            out.append((line["bbox"][1], line["bbox"][3], spans[0]["size"], text))
    doc.close()
    return sorted(out)


def oxi_lines(deck: Path, page: int) -> list[tuple[float, float, str]]:
    """(y, font_size, text) per paragraph, from --dump-layout."""
    import json

    tmp = Path(tempfile.mkdtemp(prefix="ptd_")) / "layout.json"
    subprocess.run(
        [str(EXE), str(deck), "x", "150", f"--dump-layout={tmp}"],
        capture_output=True,
        timeout=900,
    )
    data = json.loads(tmp.read_text(encoding="utf-8"))
    out = []
    for shape in data["slides"][page - 1]["shapes"]:
        content = shape.get("content")
        if not isinstance(content, dict):
            continue
        for para in content.get("paragraphs", []) or []:
            text = "".join(r.get("text", "") for r in para.get("runs", []))
            out.append((shape.get("y"), shape.get("x"), text))
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("deck")
    ap.add_argument("slide", type=int)
    ap.add_argument("--all", action="store_true", help="print every line, not just the first 30")
    args = ap.parse_args()

    deck = deck_path(args.deck)
    pdf = DEV / "pdf" / f"{deck.stem}.pdf"
    ppt = powerpoint_lines(pdf, args.slide)
    print(f"{deck.stem} slide {args.slide}: PowerPoint draws {len(ppt)} text lines")
    limit = len(ppt) if args.all else min(len(ppt), 30)
    prev = None
    for y0, y1, size, text in ppt[:limit]:
        gap = "" if prev is None else f"  d={y0 - prev:+6.2f}"
        prev = y0
        print(f"  y {y0:7.2f}..{y1:7.2f}  sz {size:5.2f}{gap}  {text[:56]!r}")

    oxi = oxi_lines(deck, args.slide)
    print(f"\nOxi lays out {len(oxi)} paragraphs on that slide (shape y / x / text):")
    for y, x, text in oxi[: (len(oxi) if args.all else 30)]:
        ys = f"{y:7.2f}" if isinstance(y, (int, float)) else str(y)
        xs = f"{x:7.2f}" if isinstance(x, (int, float)) else str(x)
        print(f"  {ys} {xs}  {text[:56]!r}")


if __name__ == "__main__":
    main()
