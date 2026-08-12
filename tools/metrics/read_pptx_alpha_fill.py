# -*- coding: utf-8 -*-
"""Export the alpha_fill probe through PowerPoint and read the composited RGB.

Prints, per arm, the colour PowerPoint painted in the middle of the translucent
rect next to the prediction of the straight source-over blend

    out = a*src + (1-a)*dst        a = alpha/100000

so the model can be accepted or rejected on the numbers.
"""
from __future__ import annotations

import sys
from pathlib import Path

import fitz
import numpy as np
import win32com.client

sys.path.insert(0, str(Path(__file__).parent))
from gen_pptx_alpha_fill import ARMS, OUT  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PPTX = OUT / "alpha_fill.pptx"
PDF = OUT / "alpha_fill.pdf"


def export() -> None:
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(PPTX), WithWindow=False)
        try:
            pres.SaveAs(str(PDF), 32)
        finally:
            pres.Close()
    finally:
        app.Quit()


def blend(src: str, dst, a: float):
    """Source-over on sRGB bytes; PowerPoint quantises alpha to 8 bits.

    The PDF carries `/ca .50196` for `<a:alpha val="50000"/>` -- that is
    round(0.5*255)/255, so the model has to quantise too.
    """
    a = round(a * 255.0) / 255.0
    s = np.array([int(src[i:i + 2], 16) for i in (0, 2, 4)], dtype=float)
    return a * s + (1.0 - a) * np.asarray(dst, dtype=float)


def main() -> None:
    if not PDF.exists() or PDF.stat().st_mtime < PPTX.stat().st_mtime:
        export()
    doc = fitz.open(PDF)
    print(f"{'arm':24s} {'PowerPoint':>16s} {'source-over':>16s} {'d':>6s}")
    worst = 0.0
    for i, (label, backdrop, fills) in enumerate(ARMS):
        pm = doc[i].get_pixmap(dpi=150)
        img = np.frombuffer(pm.samples, dtype=np.uint8)
        img = img.reshape(pm.height, pm.width, pm.n)[:, :, :3].astype(float)
        h, w = img.shape[:2]
        got = img[h // 2 - 4:h // 2 + 4, w // 2 - 4:w // 2 + 4].reshape(-1, 3)
        got = got.mean(0)
        want = np.array([255.0, 255.0, 255.0])
        if backdrop:
            want = blend(backdrop, want, 1.0)
        for hexval, alpha in fills:
            want = blend(hexval, want, alpha / 100000.0)
        d = float(np.abs(got - want).max())
        worst = max(worst, d)
        print(f"{label:24s} {str(got.round(1)):>16s} "
              f"{str(want.round(1)):>16s} {d:6.1f}")
    print(f"\nworst channel error vs source-over: {worst:.1f}")


if __name__ == "__main__":
    main()
