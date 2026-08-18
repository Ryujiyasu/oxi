# -*- coding: utf-8 -*-
"""Export the mixed-size pitch probe and solve for the ascent/descent split."""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pymupdf
import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\mixedpitch\mixedpitch.pptx").resolve()
DST = SRC.with_suffix(".pdf")
PAIRS = [(10, 10), (10, 20), (20, 10), (10, 40), (40, 10), (24, 66), (55, 66), (66, 55)]
FONTS = ["Arial", "Georgia", "Calibri", "Verdana"]
LNSPC = [("ls150_both", 150, 150), ("ls150_next", 100, 150), ("ls150_prev", 150, 100)]


# 1.2 * tmDescent / (tmAscent + tmDescent), measured with GDI at em=2048.
EXPECTED = {"Arial": 0.2276, "Georgia": 0.2315, "Calibri": 0.2640, "Verdana": 0.2073}


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


def baselines(page) -> dict[str, float]:
    ys: dict[str, float] = {}
    for blk in page.get_text("rawdict")["blocks"]:
        for ln in blk.get("lines", []):
            for sp in ln["spans"]:
                t = "".join(c["c"] for c in sp["chars"]).strip()
                if t in ("AAA", "BBB") and t not in ys:
                    ys[t] = sp["chars"][0]["origin"][1]
    return ys


def main() -> None:
    if "--noexport" not in sys.argv:
        export()
    doc = pymupdf.open(DST)
    i = 0
    for font in FONTS:
        rows = []
        print(f"\n=== {font} ===")
        print(f"{'pair':>10s} {'step pt':>9s} {'/1.2':>8s}")
        for s1, s2 in PAIRS:
            ys = baselines(doc[i]); i += 1
            step = ys["BBB"] - ys["AAA"]
            rows.append((s1, s2, step))
            print(f"{s1:4g}->{s2:<4g} {step:9.3f} {step / 1.2:8.3f}")
        # least squares for step = d*s1 + a*s2
        M = np.array([[s1, s2] for s1, s2, _ in rows], dtype=float)
        y = np.array([st for _, _, st in rows], dtype=float)
        (d, a), *_ = np.linalg.lstsq(M, y, rcond=None)
        pred = M @ np.array([d, a])
        print(f"  fit: step = {d:.4f}*prev + {a:.4f}*next   (a+d = {a + d:.4f})"
              f"  max resid {np.abs(pred - y).max():.3f}pt")
        exp = EXPECTED.get(font)
        if exp:
            print(f"  GDI tmAscent/tmDescent predicts d = {exp:.4f}"
                  f"   (fit is off by {abs(exp - d):.4f})")
    print("")
    print("=== lnSpc (Arial 10 -> 40; the 100% step is 41.16) ===")
    for label, l1, l2 in LNSPC:
        ys = baselines(doc[i]); i += 1
        print(f"  {label:12s} step {ys['BBB'] - ys['AAA']:8.3f}pt")


if __name__ == "__main__":
    main()
