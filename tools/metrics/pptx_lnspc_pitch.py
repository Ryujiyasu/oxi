# -*- coding: utf-8 -*-
"""What line pitch does PowerPoint actually use for a fixed `spcPts`?

`<a:lnSpc><a:spcPts val="3220"/>` reads as 32.20pt, and Oxi sets 32.20pt. blind
31 slide 24 is set at **32.00**: fifteen lines of body accumulate 2.88pt of drift
against the truth PDF, with the line breaks and every line width matching to
0.48pt. So PowerPoint quantises the value, and the corpus has enough distinct
fractions to say how.

Measures the pitch from the PDF itself rather than from a raster: text baselines
when the page has real text, and the placement matrices of the stencil images
when it does not (a deck whose face could not be embedded is drawn as pictures --
see `pptx_pdf_stencil_census.py`).

    python tools/metrics/pptx_lnspc_pitch.py 36 9 18
    python tools/metrics/pptx_lnspc_pitch.py 31 15 24

Prints every run of ~evenly spaced lines with its fitted pitch, so the answer can
be read beside the `spcPts` the slide declares.
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

import numpy as np
import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
PLACE = re.compile(r"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) cm\s*/(\w+) Do")


def declared(doc: str, slide: int) -> list[str]:
    import json
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    src = ROOT / "pptx" / next(i["local"] for i in manifest if f"{i['idx']:02d}" == doc)
    with zipfile.ZipFile(src) as z:
        x = z.read(f"ppt/slides/slide{slide}.xml").decode("utf-8", "replace")
    return sorted({f"{int(v)/100:.2f}" for v in re.findall(r'<a:spcPts val="(\d+)"/>', x)})


def blocks(page: pymupdf.Page) -> list[list[float]]:
    """Per text frame, its line tops -- from the text layer, else from the stencils.

    ★Grouping matters: a page mixes frames set at different `spcPts`, so a
    single sorted list of every line on the page fits a pitch that belongs to no
    shape. The PDF's own block structure separates them; on a stencilled page
    the runs are grouped by their left edge instead.
    """
    out = []
    for b in page.get_text("dict")["blocks"]:
        if b["type"] != 0 or len(b["lines"]) < 3:
            continue
        out.append(sorted(round(l["bbox"][1], 2) for l in b["lines"]))
    if out:
        return out
    rows: dict[float, list[float]] = {}
    for m in PLACE.finditer(page.read_contents().decode("latin-1")):
        x, y = round(float(m.group(5)), 2), round(-float(m.group(6)), 2)
        key = min(rows, key=lambda k: abs(k - x), default=None)
        if key is None or abs(key - x) > 6:
            key = x
        rows.setdefault(key, []).append(y)
    return [sorted(set(v)) for v in rows.values() if len(set(v)) >= 3]


def main() -> None:
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    doc = f"{int(sys.argv[1]):02d}"
    pdf = pymupdf.open(ROOT / "ssim_pptx" / "ppt_pdf" / f"{doc}.pdf")
    for arg in sys.argv[2:]:
        slide = int(arg)
        print(f"d{doc} s{slide}: spcPts declared {declared(doc, slide)}")
        for ys in blocks(pdf[slide - 1]):
            run: list[float] = []
            for y in ys:
                if run and not (2 < y - run[-1] < 140):
                    if len(run) >= 3:
                        report(run)
                    run = []
                run.append(y)
            if len(run) >= 3:
                report(run)


def report(run: list[float]) -> None:
    steps = np.diff(run)
    med = float(np.median(steps))
    keep = [run[0]]
    for y, s in zip(run[1:], steps):
        if abs(s - med) < 1.5:
            keep.append(y)
        else:
            if len(keep) >= 3:
                emit(keep)
            keep = [y]
    if len(keep) >= 3:
        emit(keep)


def emit(run: list[float]) -> None:
    slope = float(np.polyfit(range(len(run)), run, 1)[0])
    print(f"   {len(run):2d} lines  y {run[0]:8.2f} .. {run[-1]:8.2f}   pitch {slope:7.3f}pt")


if __name__ == "__main__":
    main()
