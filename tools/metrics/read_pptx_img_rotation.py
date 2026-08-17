# -*- coding: utf-8 -*-
"""Export the image-rotation probe with PowerPoint and read the corner map.

Prints, per arm, which source quadrant landed at each of the shape box's four
quadrant centres. With `--oxi` it renders the same deck with
oxi-pptx-renderer and reports agreement point by point.
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

SRC = Path(r"pipeline_data\pptx_probes\img_rotation\img_rotation.pptx").resolve()
DST = SRC.with_suffix(".pdf")
DPI = 150
BOX = (2743200 / 12700, 1600200 / 12700, 3657600 / 12700)
EXE = Path(__file__).resolve().parents[2] / "tools/oxi-pptx-renderer/target/release/oxi-pptx-renderer.exe"

LABELS = [
    "E1 pic rot=0",
    "E2 pic rot=90",
    "E3 pic rot=30",
    "E4 shape fill rot=90 rotWithShape=1",
    "E5 shape fill rot=90 rotWithShape=0",
    "E6 pic rot=90 flipH",
    "E7 pic flipH only",
    "E8 shape fill rot=30 flipH",
]
POINTS = {"TL": (0.25, 0.25), "TR": (0.75, 0.25), "BL": (0.25, 0.75), "BR": (0.75, 0.75)}


def name(rgb: tuple[int, int, int]) -> str:
    r, g, b = rgb
    if r > 200 and g > 180 and b < 100:
        return "BR"
    if r > 180 and g < 100 and b < 100:
        return "TL"
    if g > 140 and r < 120 and b < 120:
        return "TR"
    if b > 180 and r < 120 and g < 120:
        return "BL"
    if r > 230 and g > 230 and b > 230:
        return "--"
    return "??"


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
    out = Path(tempfile.mkdtemp(prefix="imgrot_oxi_"))
    subprocess.run([str(EXE), str(SRC), str(out / "slide"), str(DPI)], capture_output=True, timeout=600)
    pages = sorted(out.glob("slide_s*.png"), key=lambda p: int(p.stem.split("_s")[1]))
    return [np.asarray(Image.open(p).convert("RGB")) for p in pages]


def main() -> None:
    if "--noexport" not in sys.argv:
        export()
    with_oxi = "--oxi" in sys.argv
    pdf = pymupdf.open(DST)
    s = DPI / 72
    x0, y0, side = BOX
    oxi = oxi_pages() if with_oxi else []
    agree = mismatch = 0
    for i, label in enumerate(LABELS):
        pix = pdf[i].get_pixmap(matrix=pymupdf.Matrix(s, s), alpha=False)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        Image.fromarray(img[:, :, :3]).save(SRC.parent / f"_arm{i + 1}.png")
        cells = []
        for pname, (fx, fy) in POINTS.items():
            px, py = int(round((x0 + side * fx) * s)), int(round((y0 + side * fy) * s))
            truth = name(tuple(int(v) for v in img[py, px][:3]))
            if with_oxi and i < len(oxi):
                o = oxi[i]
                got = name(tuple(int(v) for v in o[min(py, o.shape[0] - 1), min(px, o.shape[1] - 1)][:3]))
                same = got == truth
                agree += same
                mismatch += not same
                cells.append(f"{pname}: PPT {truth} / Oxi {got}{'' if same else '  <-- MISMATCH'}")
            else:
                cells.append(f"{pname}: {truth}")
        print(f"  {label:38s} " + "   ".join(cells))
    pdf.close()
    if with_oxi:
        print(f"\nsample points: {agree} agree / {mismatch} mismatch")


if __name__ == "__main__":
    main()
