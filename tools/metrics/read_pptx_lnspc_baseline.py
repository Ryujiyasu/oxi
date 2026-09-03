# -*- coding: utf-8 -*-
"""Export `lnspc_baseline.pptx` and read the baseline PowerPoint gave each arm.

Exports the probe from its OWN PowerPoint session (`pptx_truth_pdf_first_open_is_cold`),
then reads every span's origin out of the PDF -- the origin IS the baseline, so
this needs no pixels and no thresholds.

Reports, per arm, the offset from the shape's top to the first baseline and the
step to the second, both in points and in ems, beside what the engine's
`first_baseline_off` rule would give:

    quarter rule   descent = natural - 0.25 * 1.2 * fs * (1 - n)   (implemented)
    scaled rule    descent = natural * n                            (deck 40's reading)

    python tools/metrics/read_pptx_lnspc_baseline.py
"""
from __future__ import annotations

import os
import re
import sys

SRC = os.path.abspath(os.path.join("tools", "metrics", "lnspc_baseline.pptx"))
PDF = os.path.abspath(os.path.join("tools", "metrics", "lnspc_baseline.pdf"))
FACES = ["Arial", "Georgia", "Verdana", "Calibri",
         "Segoe Script", "Papyrus", "Viner Hand ITC", "Javanese Text"]
SIZES = [24, 60]
SPACINGS = [0.4, 0.6, 0.8, 1.0, 1.2, 1.5]


def export() -> None:
    import win32com.client

    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(SRC, WithWindow=False)
    if os.path.exists(PDF):
        os.remove(PDF)
    pres.SaveAs(PDF, 32)
    pres.Close()
    app.Quit()


def main() -> None:
    import pymupdf

    if "--keep" not in sys.argv or not os.path.exists(PDF):
        export()
    pdf = pymupdf.open(PDF)
    print("%-9s %5s %5s   %8s %8s   %8s   %8s %8s"
          % ("face", "size", "lnSpc", "top->b1", "b1 em", "step", "quarter", "scaled"))
    rows = []
    for pno in range(len(pdf)):
        pg = pdf[pno]
        face, size = FACES[pno // len(SIZES)], SIZES[pno % len(SIZES)]
        # Every span on the page, grouped by the shape it came from: the shapes
        # sit on a 3 x 2 grid, so the shape is recoverable from the origin.
        spans = []
        for blk in pg.get_text("dict")["blocks"]:
            for ln in blk.get("lines", []):
                for sp in ln["spans"]:
                    if sp["text"].strip():
                        spans.append((sp["origin"][0], sp["origin"][1], sp["size"]))
        # shape tops, in points, from the generator's own EMU grid
        for i, mult in enumerate(SPACINGS):
            sx = (200000 + (i % 3) * 2900000) / 12700.0
            sy = (200000 + (i // 3) * 2300000) / 12700.0
            mine = sorted((s for s in spans if abs(s[0] - sx) < 12 and sy - 4 < s[1] < sy + 260),
                          key=lambda s: s[1])
            if len(mine) < 2:
                print("%-9s %5d %5.1f   (found %d spans)" % (face, size, mult, len(mine)))
                continue
            b1 = mine[0][1] - sy
            step = mine[1][1] - mine[0][1]
            pitch = size * 1.2
            quarter = 0.25 * pitch
            # The engine's own numbers need the face's ascent split, which the
            # probe recovers from the n = 1 arm: there both rules agree.
            rows.append((face, size, mult, b1, step))
            print("%-9s %5d %5.1f   %8.2f %8.4f   %8.2f" % (face, size, mult, b1, b1 / size, step))
    # With the n = 1 arm as the face's natural ascent, both candidate rules are
    # predictions rather than fits.
    print()
    print("%-9s %5s %5s   %8s %8s %8s   %s"
          % ("face", "size", "lnSpc", "measured", "quarter", "scaled", "which"))
    natural = {(f, s): b for f, s, m, b, _ in rows if abs(m - 1.0) < 1e-6}
    for face, size, mult, b1, step in rows:
        asc = natural.get((face, size))
        if asc is None:
            continue
        pitch = size * 1.2
        nat_desc = pitch - asc
        quarter = 0.25 * pitch
        if mult <= 1.0:
            desc_q = max(nat_desc + quarter * (mult - 1.0), min(nat_desc, quarter * mult))
        else:
            desc_q = max(nat_desc, quarter * mult)
        pred_q = pitch * mult - desc_q
        pred_s = asc * mult
        which = "quarter" if abs(b1 - pred_q) < abs(b1 - pred_s) else "scaled"
        if abs(pred_q - pred_s) < 0.05:
            which = "-"
        print("%-9s %5d %5.1f   %8.2f %8.2f %8.2f   %s"
              % (face, size, mult, b1, pred_q, pred_s, which))


if __name__ == "__main__":
    main()
