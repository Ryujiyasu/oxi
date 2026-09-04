# -*- coding: utf-8 -*-
"""In which colour space does PowerPoint interpolate a gradient?

Its PDF export answers directly: every two-stop gradient becomes a
`/FunctionType 0` sampled ramp (256 x RGB, 8-bit), so the interpolation curve
is written down, endpoint to endpoint. For each such table this fits two
models between the TABLE's own endpoints -- straight sRGB lerp, and lerp in
linear light (sRGB decode, lerp, encode) -- and reports which explains the 254
interior samples.

d24's accent4->accent3 ramp picked linear light to 1/255 on every probed
sample; this asks the whole corpus.

    python tools/metrics/pptx_gradient_space_census.py [--blind]
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import numpy as np
import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[2] / "pipeline_data" / "pptx_benchmark"


def srgb_decode(v: np.ndarray) -> np.ndarray:
    v = v / 255.0
    return np.where(v <= 0.04045, v / 12.92, ((v + 0.055) / 1.055) ** 2.4)


def srgb_encode(x: np.ndarray) -> np.ndarray:
    return np.where(x <= 0.0031308, x * 12.92, 1.055 * x ** (1 / 2.4) - 0.055) * 255.0


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--blind", action="store_true")
    args = ap.parse_args()
    pdfs = sorted((ROOT / ("ssim_pptx/ppt_pdf" if args.blind else "dev/pdf")).glob("*.pdf"))

    ramps = 0
    lin_wins = 0
    srgb_wins = 0
    ties = 0
    worst = []
    for path in pdfs:
        try:
            doc = pymupdf.open(path)
        except Exception:
            continue
        for xref in range(1, doc.xref_length()):
            try:
                obj = doc.xref_object(xref)
            except Exception:
                continue
            if "/FunctionType 0" not in obj or "/BitsPerSample 8" not in obj:
                continue
            m = re.search(r"/Size \[ (\d+) \]", obj)
            if not m or m.group(1) != "256":
                continue
            try:
                arr = np.frombuffer(doc.xref_stream(xref), np.uint8)
            except Exception:
                continue
            if arr.size != 256 * 3:
                continue
            table = arr.reshape(256, 3).astype(np.float64)
            c0, c1 = table[0], table[-1]
            if np.abs(c0 - c1).max() < 8:
                continue  # too flat to distinguish anything
            t = np.linspace(0, 1, 256)[:, None]
            lerp = c0 + (c1 - c0) * t
            lin = srgb_encode(srgb_decode(c0) + (srgb_decode(c1) - srgb_decode(c0)) * t)
            e_lerp = np.abs(lerp - table).max()
            e_lin = np.abs(lin - table).max()
            ramps += 1
            if e_lin + 1.0 < e_lerp:
                lin_wins += 1
            elif e_lerp + 1.0 < e_lin:
                srgb_wins += 1
            else:
                ties += 1
            worst.append((min(e_lin, e_lerp), path.stem[:20], xref, e_lin, e_lerp))
        doc.close()
    print("%d two-endpoint 256-sample ramps: linear-light explains %d, "
          "sRGB-lerp %d, indistinguishable %d" % (ramps, lin_wins, srgb_wins, ties))
    worst.sort(reverse=True)
    print("\nleast-well-explained (residual of the better model):")
    for r, stem, xref, el, es in worst[:8]:
        print("  %-22s xref %-5d best %.1f  (linear %.1f / srgb %.1f)"
              % (stem, xref, r, el, es))


if __name__ == "__main__":
    main()
