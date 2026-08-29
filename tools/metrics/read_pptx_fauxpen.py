# -*- coding: utf-8 -*-
"""Read the faux-pen probe: is Oxi's synthesised bold as heavy as PowerPoint's?

Renders both arms with the renderer as it stands (and again with
`OXI_FAUXPEN_DISABLE=1` when `--off` is given) and reports the INK AREA of the
text on each slide -- the integral of coverage, which counts an antialiased edge
once instead of cutting it at some grey.

The number that matters is the last column. Ink area on its own carries the
constant bias between two rasterisers; the BOLD-minus-PLAIN difference does not,
because the same bias sits in both terms. That difference is the pen:

    PowerPoint  bold - plain      what `size/35` is worth
    Oxi         bold - plain      what the dilation actually laid down

Usage:
    python tools/metrics/read_pptx_fauxpen.py [--off] [--dpi 150]
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_probes" / "fauxpen"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"


def render(arm: str, dpi: int, env_extra: dict[str, str]) -> Path:
    d = Path(tempfile.mkdtemp(prefix=f"fauxpen_{arm}_"))
    env = dict(os.environ)
    env.update(env_extra)
    subprocess.run([str(EXE), str(OUT / f"{arm}.pptx"), str(d / "slide"), str(dpi)],
                   capture_output=True, env=env, timeout=1800, check=False)
    return d


def ink(a: np.ndarray, dpi: int) -> float:
    """Coverage integral in pt^2, on a white ground."""
    k = dpi / 72.0
    return float((1.0 - a / 255.0).sum()) / (k * k)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--off", action="store_true",
                    help="also render with OXI_FAUXPEN_DISABLE=1, as the before arm")
    args = ap.parse_args()

    arms = {a: render(a, args.dpi, {}) for a in ("plain", "bold")}
    offs = ({a: render(a, args.dpi, {"OXI_FAUXPEN_DISABLE": "1"})
             for a in ("plain", "bold")} if args.off else {})

    k = args.dpi / 72.0
    refs = {a: pymupdf.open(OUT / f"{a}.pdf") for a in ("plain", "bold")}
    n = len(refs["bold"])
    head = f"{'size':>6}{'PPT plain':>11}{'PPT bold':>10}{'PPT pen':>9}"
    head += f"{'oxi plain':>11}{'oxi bold':>10}{'oxi pen':>9}{'pen err':>9}"
    if args.off:
        head += f"{'OFF pen':>9}"
    print(head)
    for i in range(n):
        row = {}
        for a in ("plain", "bold"):
            pix = refs[a][i].get_pixmap(matrix=pymupdf.Matrix(k, k))
            ref = np.asarray(Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                             .convert("L"), float)
            row[f"ref_{a}"] = ink(ref, args.dpi)
            img = Image.open(arms[a] / f"slide_s{i + 1}.png").convert("L")
            if img.size != (pix.width, pix.height):
                img = img.resize((pix.width, pix.height), Image.LANCZOS)
            row[f"oxi_{a}"] = ink(np.asarray(img, float), args.dpi)
            if offs:
                im2 = Image.open(offs[a] / f"slide_s{i + 1}.png").convert("L")
                if im2.size != (pix.width, pix.height):
                    im2 = im2.resize((pix.width, pix.height), Image.LANCZOS)
                row[f"off_{a}"] = ink(np.asarray(im2, float), args.dpi)
        size = [12, 24, 48, 96][i] if i < 4 else i + 1
        rp = row["ref_bold"] - row["ref_plain"]
        op = row["oxi_bold"] - row["oxi_plain"]
        line = (f"{size:>6}{row['ref_plain']:>11.1f}{row['ref_bold']:>10.1f}{rp:>9.1f}"
                f"{row['oxi_plain']:>11.1f}{row['oxi_bold']:>10.1f}{op:>9.1f}"
                f"{100 * (op - rp) / rp if rp else 0:>8.1f}%")
        if offs:
            fp = row["off_bold"] - row["off_plain"]
            line += f"{fp:>9.1f}"
        print(line)
    for d in refs.values():
        d.close()


if __name__ == "__main__":
    main()
