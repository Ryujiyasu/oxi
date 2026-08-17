# -*- coding: utf-8 -*-
"""Export the shape-blipFill probe with PowerPoint and read where the ink went.

Prints, per arm, the colour PowerPoint put at a grid of points across (and
outside) the shape box, named by which source quadrant that colour is:
red = source top-left, green = top-right, blue = bottom-left, yellow =
bottom-right, white = nothing drawn there.
"""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pymupdf
import win32com.client
from PIL import Image

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\custgeom_blipfill\custgeom_blipfill.pptx").resolve()
DST = SRC.with_suffix(".pdf")
DPI = 150
BOX = (2743200 / 12700, 1600200 / 12700, 3657600 / 12700)  # x, y, side (pt)

LABELS = [
    "D1 rect, fillRect 0",
    "D2 triangle, fillRect 0",
    "D3 rect, fillRect r=-100%",
    "D4 triangle, fillRect r=-100% b=-100%",
    "D5 rect, srcRect l=25% r=25%",
    "D6 rect, the d28 fillRect",
]

# Sampled as fractions of the shape box; negatives / >1 probe OUTSIDE it.
POINTS = {
    "TL in": (0.25, 0.25),
    "TR in": (0.75, 0.25),
    "BL in": (0.25, 0.75),
    "BR in": (0.75, 0.75),
    "corner TL": (0.05, 0.05),
    "corner TR": (0.95, 0.05),
    "right out": (1.30, 0.50),
    "below out": (0.50, 1.30),
}


def name(rgb: tuple[int, int, int]) -> str:
    r, g, b = rgb
    if r > 200 and g > 180 and b < 100:
        return "YELLOW(src BR)"
    if r > 180 and g < 100 and b < 100:
        return "RED(src TL)"
    if g > 140 and r < 120 and b < 120:
        return "GREEN(src TR)"
    if b > 180 and r < 120 and g < 120:
        return "BLUE(src BL)"
    if r > 230 and g > 230 and b > 230:
        return "white(none)"
    if r < 60 and g < 60 and b < 60:
        return "black(frame)"
    return f"other{rgb}"


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
    print("exported", DST, DST.stat().st_size, "bytes")


def oxi_pages() -> list[np.ndarray]:
    """Render the probe with oxi-pptx-renderer and return its pages."""
    import subprocess
    import tempfile

    exe = Path(__file__).resolve().parents[2] / "tools/oxi-pptx-renderer/target/release/oxi-pptx-renderer.exe"
    out = Path(tempfile.mkdtemp(prefix="blipfill_oxi_"))
    subprocess.run([str(exe), str(SRC), str(out / "slide"), str(DPI)], capture_output=True, timeout=600)
    pages = sorted(out.glob("slide_s*.png"), key=lambda p: int(p.stem.split("_s")[1]))
    return [np.asarray(Image.open(p).convert("RGB")) for p in pages]


def main() -> None:
    if "--noexport" not in sys.argv:
        export()
    compare_oxi = "--oxi" in sys.argv
    pdf = pymupdf.open(DST)
    s = DPI / 72
    x0, y0, side = BOX
    oxi = oxi_pages() if compare_oxi else []
    agree = disagree = 0
    for i, label in enumerate(LABELS):
        pix = pdf[i].get_pixmap(matrix=pymupdf.Matrix(s, s), alpha=False)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        Image.fromarray(img[:, :, :3]).save(SRC.parent / f"_arm{i + 1}.png")
        print(f"  {label}")
        for pname, (fx, fy) in POINTS.items():
            px = int(round((x0 + side * fx) * s))
            py = int(round((y0 + side * fy) * s))
            if not (0 <= px < img.shape[1] and 0 <= py < img.shape[0]):
                print(f"      {pname:10s} off-page")
                continue
            truth = name(tuple(int(v) for v in img[py, px][:3]))
            if not compare_oxi or i >= len(oxi):
                print(f"      {pname:10s} {truth}")
                continue
            o = oxi[i]
            oy, ox = min(py, o.shape[0] - 1), min(px, o.shape[1] - 1)
            got = name(tuple(int(v) for v in o[oy, ox][:3]))
            same = got == truth
            agree += same
            disagree += not same
            print(f"      {pname:10s} PowerPoint {truth:14s} Oxi {got:14s} {'ok' if same else 'MISMATCH'}")
    pdf.close()
    if compare_oxi:
        print(f"\nsample points: {agree} agree / {disagree} mismatch")


if __name__ == "__main__":
    main()
