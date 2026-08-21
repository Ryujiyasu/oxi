# -*- coding: utf-8 -*-
"""Read the cellctr probe: is a centred cell's baseline offset constant?

Per row, prints the spare room (row height minus 2*margin minus 1.2*size) and
the offset of PowerPoint's baseline below the centred block top. A constant
column means centring is not the variable.

Usage: python tools/metrics/read_pptx_cellctr.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DIR = Path(r"pipeline_data\pptx_probes\cellctr").resolve()


def main() -> None:
    manifest = json.loads((DIR / "cellctr_manifest.json").read_text(encoding="utf-8"))
    pdf = pymupdf.open(DIR / "deck.pdf")
    for row in manifest:
        page = pdf[row["slide"] - 1]
        ys = []
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for line in b["lines"]:
                for s in line["spans"]:
                    ch = s["chars"]
                    if "".join(c["c"] for c in ch).strip() == "Hxy":
                        ys.append(ch[0]["origin"][1])
        ys.sort()
        fs, mar = row["size"], row["margin"]
        block = 1.2 * fs
        print(f"\n== {row['face']} {fs}pt margin {mar}")
        print(f"   {'row h':>7s} {'spare':>7s} {'baseline':>9s} {'A (ctr)':>9s} "
              f"{'A (top)':>9s}")
        top = row["table_top"]
        for h, y in zip(row["heights"], ys):
            inner_h = h - 2 * mar
            spare = inner_h - block
            ctr_top = top + mar + max(spare / 2, 0.0)
            print(f"   {h:7.1f} {spare:7.2f} {y:9.2f} "
                  f"{(y - ctr_top) / fs:9.4f} {(y - top - mar) / fs:9.4f}")
            top += h
        if len(ys) < len(row["heights"]):
            print(f"   (only {len(ys)} baselines found)")


if __name__ == "__main__":
    main()
