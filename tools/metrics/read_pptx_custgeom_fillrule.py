# -*- coding: utf-8 -*-
"""Export the custGeom fill-rule probe with PowerPoint and read the verdict.

Samples the centre of each arm's shape plus a ring point that both rules agree
on, so a null result (nothing drawn at all) cannot be mistaken for "hole".
"""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pymupdf
import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\custgeom_fillrule\custgeom_fillrule.pptx").resolve()
DST = SRC.with_suffix(".pdf")
DPI = 150
# The shape box in points (EMU/12700), from the generator.
BOX = (2743200 / 12700, 1600200 / 12700, 3657600 / 12700)

# (label, probe point, control point, what each rule predicts at the probe
# point). Both points are fractions of the shape box; the control is a place
# both rules fill, so "nothing was drawn at all" cannot read as "hole".
EXPECT = [
    ("C1 nested, same winding", (0.5, 0.5), (0.08, 0.5), "even-odd HOLE / nonzero FILLED"),
    ("C2 nested, opposite winding", (0.5, 0.5), (0.08, 0.5), "both HOLE"),
    ("C3 pentagram", (0.5, 0.5), (0.5, 0.15), "even-odd HOLE / nonzero FILLED"),
    ("C4 two disjoint squares", (0.2, 0.2), (0.8, 0.8), "both FILLED (sanity)"),
    ("C5 three nested, same winding", (0.5, 0.5), (0.08, 0.5), "both FILLED"),
]


def export() -> None:
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        prs = app.Presentations.Open(str(SRC), WithWindow=False)
        try:
            prs.SaveAs(str(DST), 32)  # ppSaveAsPDF
        finally:
            prs.Close()
    finally:
        app.Quit()
    print("exported", DST, DST.stat().st_size, "bytes")


def main() -> None:
    if "--noexport" not in sys.argv:
        export()
    pdf = pymupdf.open(DST)
    s = DPI / 72
    x0, y0, side = BOX
    for i, (label, probe, control, expect) in enumerate(EXPECT):
        pix = pdf[i].get_pixmap(matrix=pymupdf.Matrix(s, s), alpha=False)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)

        def at(pt: tuple[float, float]) -> tuple[int, int, int]:
            px = int((x0 + side * pt[0]) * s)
            py = int((y0 + side * pt[1]) * s)
            return tuple(int(v) for v in img[py, px][:3])

        p, c = at(probe), at(control)
        red = lambda v: v[0] > 128 > v[1]
        print(
            f"  {label:34s} probe={p} {'FILLED' if red(p) else 'HOLE  '}"
            f"  control={c} {'ok' if red(c) else 'NOT DRAWN'}   [{expect}]"
        )
    pdf.close()


if __name__ == "__main__":
    main()
