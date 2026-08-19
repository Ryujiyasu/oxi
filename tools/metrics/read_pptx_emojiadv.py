# -*- coding: utf-8 -*-
"""Read the bracketed-emoji advances out of PowerPoint's PDF."""
from __future__ import annotations

import sys
from pathlib import Path

import pymupdf
import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\emojiadv\emojiadv.pptx").resolve()
DST = SRC.with_suffix(".pdf")
ARMS = ["none", "heart_text", "heart_color", "hand", "thermo_text",
        "thermo_color", "grin", "watch"]
SIZE = 40.0


def export() -> None:
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        prs = app.Presentations.Open(str(SRC), WithWindow=False)
        try:
            prs.SaveAs(str(DST), 32)
        finally:
            prs.Close()
    finally:
        app.Quit()


def main() -> None:
    if "--noexport" not in sys.argv:
        export()
    doc = pymupdf.open(DST)
    base = None
    for i, label in enumerate(ARMS):
        xs = []
        for blk in doc[i].get_text("rawdict")["blocks"]:
            for ln in blk.get("lines", []):
                for sp in ln["spans"]:
                    for c in sp["chars"]:
                        if c["c"] == "A" and c["bbox"][1] > 60:
                            xs.append(c["bbox"][0])
        xs.sort()
        if len(xs) != 2:
            print(f"{label:14s} (expected two letters, saw {len(xs)})")
            continue
        span = xs[1] - xs[0]
        if label == "none":
            base = span
            print(f"{label:14s} letter advance {span:7.3f}pt")
            continue
        adv = span - base
        print(f"{label:14s} emoji advance {adv:7.3f}pt = {adv / SIZE:6.4f} em")


if __name__ == "__main__":
    main()
