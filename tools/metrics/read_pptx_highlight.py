# -*- coding: utf-8 -*-
"""Read the highlight rectangle and the line it sits on out of PowerPoint."""
from __future__ import annotations

import sys
from pathlib import Path

import pymupdf
import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\highlight\highlight.pptx").resolve()
DST = SRC.with_suffix(".pdf")
ARMS = [("mid18", 18), ("mid36", 36), ("mid18_ln150", 18), ("mid18_ln70", 18),
        ("start18", 18), ("end18", 18), ("space18", 18), ("desc18", 18),
        ("arial96", 96), ("times18", 18), ("times96", 96), ("courier18", 18),
        ("segoe18", 18), ("georgia96", 96), ("tallneighbour", 18),
        ("bahn96", 96), ("cascadia96", 96)]


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
    for i, (label, sz) in enumerate(ARMS):
        page = doc[i]
        rects = [d["rect"] for d in page.get_drawings()
                 if d.get("fill") and d["rect"].y0 > 60 and d["rect"].width > 5]
        chars = []
        for blk in page.get_text("rawdict")["blocks"]:
            for ln in blk.get("lines", []):
                for sp in ln["spans"]:
                    if sp["bbox"][1] < 60:
                        continue
                    for c in sp["chars"]:
                        chars.append((c["c"], c["origin"][0], c["origin"][1], c["bbox"]))
        if not rects or not chars:
            print(f"{label:14s} rects={len(rects)} chars={len(chars)}")
            continue
        r = rects[0]
        base = chars[0][2]
        print(f"{label:14s} sz={sz:2d}  rect x {r.x0:7.2f}..{r.x1:7.2f}  "
              f"y {r.y0:7.2f}..{r.y1:7.2f}  h={r.height:6.2f} ({r.height / sz:5.3f} em)  "
              f"baseline {base:7.2f}  top-above-baseline {base - r.y0:6.2f} "
              f"({(base - r.y0) / sz:5.3f} em)  bottom {r.y1 - base:5.2f} "
              f"({(r.y1 - base) / sz:5.3f} em)")
        text = "".join(c[0] for c in chars)
        print(f"{'':14s} text={text!r}")
        for c in chars:
            if abs(c[1] - r.x0) < 2 or abs(c[3][2] - r.x1) < 2:
                print(f"{'':14s} edge char {c[0]!r} pen={c[1]:.2f} right={c[3][2]:.2f}")


if __name__ == "__main__":
    main()
