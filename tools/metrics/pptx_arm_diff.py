# -*- coding: utf-8 -*-
"""One slide of the blind corpus, in both arms of a flag, beside PowerPoint.

`pptx_flag_ab.py` says a deck moved and by how much; it does not say WHAT moved,
and a mean that improves while one deck's worst slide drops is exactly the case
where the number is not the answer. This renders the one slide twice -- feature
off, feature on -- and writes the three panels the eye needs:

    PowerPoint     Oxi (feature ON)     Oxi vs PowerPoint     where the ARMS differ

The first three are the usual three panels. The fourth is the arm-to-arm
difference, which the usual three cannot show: it is what THIS CHANGE did, with
everything the change did not touch subtracted away.

★The renderer is NOT parallel-safe (`pptx_render_not_parallel_safe`), so this
must not run while `pptx_flag_ab.py` does.

    python tools/metrics/pptx_arm_diff.py OXI_MUDRAW_DISABLE 21 2
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
SSIMD = ROOT / "ssim_pptx"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"
DPI = 150


def render(src: Path, out: Path, flag: str, on: bool) -> None:
    """The same arm convention as `pptx_flag_ab.py`: `on` selects the FEATURE."""
    shutil.rmtree(out, ignore_errors=True)
    out.mkdir(parents=True)
    env = dict(os.environ)
    env.pop(flag, None)
    if (not on) if flag.endswith("_DISABLE") else on:
        env[flag] = "1"
    subprocess.run([str(EXE), str(src), str(out / "slide"), str(DPI)],
                   capture_output=True, env=env, timeout=3600, check=False)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("flag")
    ap.add_argument("doc", type=int)
    ap.add_argument("slide", type=int)
    ap.add_argument("--out", default="")
    args = ap.parse_args()

    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    item = next(i for i in manifest if i["idx"] == args.doc)
    src = ROOT / "pptx" / item["local"]
    off_dir, on_dir = SSIMD / "_armdiff_off", SSIMD / "_armdiff_on"
    print(f"rendering {item['local'][:50]} ...", flush=True)
    render(src, off_dir, args.flag, on=False)
    render(src, on_dir, args.flag, on=True)

    name = f"slide_s{args.slide}.png"
    a = Image.open(off_dir / name).convert("RGB")
    b = Image.open(on_dir / name).convert("RGB")
    pdf = pymupdf.open(SSIMD / "ppt_pdf" / f"{args.doc:02d}.pdf")
    pix = pdf[args.slide - 1].get_pixmap(dpi=DPI)
    truth = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    pdf.close()
    if b.size != truth.size:
        truth = truth.resize(b.size, Image.LANCZOS)
        a = a.resize(b.size, Image.LANCZOS)

    # Where the arms differ, as ink on white -- a shifted glyph shows as its own
    # doubled outline, which is what tells a MOVE from a REDRAW.
    d = np.abs(np.asarray(a).astype(int) - np.asarray(b).astype(int)).sum(axis=2)
    heat = np.full(d.shape + (3,), 255, dtype=np.uint8)
    hit = d > 12
    heat[hit] = np.array([200, 30, 30], dtype=np.uint8)
    print(f"pixels the two arms disagree on: {int(hit.sum())} "
          f"({100.0 * hit.mean():.3f}% of the slide)")

    # Oxi against the truth, the usual third panel.
    dt = np.abs(np.asarray(truth).astype(int) - np.asarray(b).astype(int)).sum(axis=2)
    tdiff = np.full(dt.shape + (3,), 255, dtype=np.uint8)
    tdiff[dt > 12] = np.array([30, 60, 200], dtype=np.uint8)

    panels = [truth, b, Image.fromarray(tdiff), Image.fromarray(heat)]
    w = sum(p.width for p in panels) + 12 * (len(panels) - 1)
    canvas = Image.new("RGB", (w, max(p.height for p in panels)), (255, 255, 255))
    x = 0
    for p in panels:
        canvas.paste(p, (x, 0))
        x += p.width + 12
    out = Path(args.out) if args.out else (
        SSIMD / f"_armdiff_{args.doc:02d}_s{args.slide}.png")
    canvas.save(out)
    print(f"wrote {out}  (PowerPoint | feature ON | vs PowerPoint | arms differ)")


if __name__ == "__main__":
    main()
