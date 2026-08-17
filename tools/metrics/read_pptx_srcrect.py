# -*- coding: utf-8 -*-
"""Export the srcRect probe with PowerPoint and read the source->box mapping.

For each arm it reports, at nine points down the shape box, which tenth of the
SOURCE image landed there -- read off the black rules the generator draws every
10% of the source height. `--oxi` renders the same deck with
oxi-pptx-renderer and prints the two side by side.
"""
from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\srcrect\srcrect.pptx").resolve()
DST = SRC.with_suffix(".pdf")
DPI = 150
BOX = (2743200 / 12700, 1600200 / 12700, 3657600 / 12700)
EXE = Path(__file__).resolve().parents[2] / "tools/oxi-pptx-renderer/target/release/oxi-pptx-renderer.exe"

LABELS = [
    "F1 no srcRect",
    "F2 srcRect b=50%",
    "F3 srcRect t=50%",
    "F4 srcRect l=50%",
    "F5 srcRect b=14.368% (d30)",
    "F6 srcRect t=25% b=25%",
]


def rules(img: np.ndarray, x0: float, y0: float, side: float, s: float) -> list[float]:
    """Vertical positions (as a fraction of the box) of the black rules."""
    col = int((x0 + side * 0.5) * s)
    top, bot = int(y0 * s), int((y0 + side) * s)
    # Near-black only: the mean of a pure red (255,0,0) is 85, so a mean
    # threshold marks the whole red quadrant as a rule.
    strip = img[top:bot, col - 2 : col + 3, :3].max(axis=2).mean(axis=1)
    dark = strip < 60
    out, run = [], []
    for i, d in enumerate(dark):
        if d:
            run.append(i)
        elif run:
            out.append((run[0] + run[-1]) / 2 / len(strip))
            run = []
    if run:
        out.append((run[0] + run[-1]) / 2 / len(strip))
    return out


def export() -> None:
    import win32com.client

    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        prs = app.Presentations.Open(str(SRC), WithWindow=False)
        try:
            prs.SaveAs(str(DST), 32)
        finally:
            prs.Close()
    finally:
        app.Quit()
    print("exported", DST, DST.stat().st_size, "bytes")


def oxi_pages() -> list[np.ndarray]:
    out = Path(tempfile.mkdtemp(prefix="srcrect_oxi_"))
    subprocess.run([str(EXE), str(SRC), str(out / "slide"), str(DPI)], capture_output=True, timeout=600)
    return [
        np.asarray(Image.open(p).convert("RGB"))
        for p in sorted(out.glob("slide_s*.png"), key=lambda p: int(p.stem.split("_s")[1]))
    ]


def main() -> None:
    if "--noexport" not in sys.argv:
        export()
    with_oxi = "--oxi" in sys.argv
    pdf = pymupdf.open(DST)
    s = DPI / 72
    x0, y0, side = BOX
    oxi = oxi_pages() if with_oxi else []
    for i, label in enumerate(LABELS):
        pix = pdf[i].get_pixmap(matrix=pymupdf.Matrix(s, s), alpha=False)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3]
        Image.fromarray(img).save(SRC.parent / f"_arm{i + 1}.png")
        p = rules(img, x0, y0, side, s)
        line = f"  {label:28s} PPT rules at " + ", ".join(f"{v:.3f}" for v in p)
        if with_oxi and i < len(oxi):
            o = rules(oxi[i], x0, y0, side, s)
            line += "\n" + " " * 32 + "Oxi rules at " + ", ".join(f"{v:.3f}" for v in o)
        print(line)
    pdf.close()


if __name__ == "__main__":
    main()
