# -*- coding: utf-8 -*-
"""Read the wrapfit sweep truth: per arm, where does PowerPoint stop breaking?

For every slide of pipeline_data/pptx_probes/wrapfit/deck.pdf, count the text
line bands (distinct baseline ys). The manifest says which arm and box width
each slide carries; the report gives per arm the largest 2-band width and the
smallest 1-band width -- the bracket around PowerPoint's fit threshold -- and
its offset from the string's design advance sum. Non-monotone flips are
flagged loudly: they would mean the probe (or the rule) is not width-only.

Usage: python tools/metrics/read_pptx_wrapfit.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DIR = Path(r"pipeline_data\pptx_probes\wrapfit").resolve()


def bands(page: pymupdf.Page) -> int:
    ys = set()
    for block in page.get_text("rawdict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            text = "".join(c["c"] for s in line["spans"] for c in s["chars"]).strip()
            if text:
                ys.add(round(line["spans"][0]["chars"][0]["origin"][1], 1))
    return len(ys)


def main() -> None:
    # Optional round suffix: `read_pptx_wrapfit.py 2` reads wrapfit2_manifest
    # + r2/deck.pdf (measure_pptx_word.py always names its output deck.pdf,
    # so each round exports into its own subdirectory).
    suffix = sys.argv[1] if len(sys.argv) > 1 else ""
    manifest = json.loads(
        (DIR / f"wrapfit{suffix}_manifest.json").read_text(encoding="utf-8"))
    pdf_path = DIR / f"r{suffix}" / "deck.pdf" if suffix else DIR / "deck.pdf"
    pdf = pymupdf.open(pdf_path)
    if pdf.page_count != len(manifest):
        print(f"WARNING: {pdf.page_count} pages vs {len(manifest)} manifest rows")
    arms: dict[str, list[dict]] = {}
    for row in manifest:
        if row["slide"] > pdf.page_count:
            break
        row["bands"] = bands(pdf[row["slide"] - 1])
        arms.setdefault(row["arm"], []).append(row)
    for arm, rows in arms.items():
        rows.sort(key=lambda r: r["width_pt"])
        s = rows[0]["design_sum_pt"]
        broken = [r for r in rows if r["bands"] > 1]
        whole = [r for r in rows if r["bands"] == 1]
        top_break = max((r["width_pt"] for r in broken), default=None)
        low_whole = min((r["width_pt"] for r in whole), default=None)
        flips = sum(
            1 for a, b in zip(rows, rows[1:])
            if a["bands"] == 1 and b["bands"] > 1
        )
        print(f"== {arm}  text={rows[0]['text']!r}  design={s}pt")
        print(f"   widest 2-band  W = {top_break}   (delta {None if top_break is None else round(top_break - s, 3)})")
        print(f"   narrowest 1-band W = {low_whole} (delta {None if low_whole is None else round(low_whole - s, 3)})")
        if flips:
            print(f"   *** NON-MONOTONE: {flips} whole->broken flips going wider ***")
        marks = "".join("2" if r["bands"] > 1 else ("1" if r["bands"] == 1 else "0") for r in rows)
        print(f"   sweep: {marks}")


if __name__ == "__main__":
    main()
