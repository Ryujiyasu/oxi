# -*- coding: utf-8 -*-
"""Declared `spcPts` against the pitch PowerPoint actually rendered, per SHAPE.

A slide mixes frames set at different fixed line spacings, so a pitch measured
over a whole page belongs to no shape. This matches each `<p:sp>` to the PDF
block that lands in its box, and prints the value the file declares beside the
pitch the PDF shows.

    python tools/metrics/pptx_lnspc_law.py 31 36 37

Why: blind 31 s24 declares `spcPts val="3220"` (32.20pt) and PowerPoint renders
**32.000**; Oxi renders 32.20 and drifts 2.88pt over fifteen lines with every
line break and width matching. Integer values are honoured exactly (a 105.00pt
frame measures 105.005), so the question is only what PowerPoint does with the
fraction -- floor, round, or round to a half point -- and the corpus carries
enough distinct fractions (33.59, 32.62, 29.40, 22.39, 37.79 ...) to separate
them.
"""
from __future__ import annotations

import json
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
EMU = 12700.0
PLACE = re.compile(r"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) cm\s*/(\w+) Do")


def shapes(xml: str) -> list[dict]:
    out = []
    for m in re.finditer(r"<p:sp>.*?</p:sp>", xml, re.S):
        s = m.group(0)
        off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', s)
        ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', s)
        spc = re.findall(r'<a:lnSpc><a:spcPts val="(\d+)"/></a:lnSpc>', s)
        if not (off and ext and spc):
            continue
        vals = {int(v) / 100 for v in spc}
        if len(vals) != 1:
            continue
        out.append({
            "x": int(off.group(1)) / EMU, "y": int(off.group(2)) / EMU,
            "w": int(ext.group(1)) / EMU, "h": int(ext.group(2)) / EMU,
            "declared": vals.pop(),
            "text": "".join(re.findall(r"<a:t>(.*?)</a:t>", s, re.S))[:28],
        })
    return out


def pdf_blocks(page: pymupdf.Page) -> list[tuple[tuple, list[float]]]:
    """(bbox, line tops) per text frame; stencil placements when text is hidden."""
    out = []
    for b in page.get_text("dict")["blocks"]:
        if b["type"] == 0 and len(b["lines"]) >= 3:
            out.append((b["bbox"], sorted(round(l["bbox"][1], 2) for l in b["lines"])))
    if out:
        return out
    rows: dict[float, list[tuple[float, float]]] = {}
    height = page.rect.height
    for m in PLACE.finditer(page.read_contents().decode("latin-1")):
        # PDF space is bottom-up: e is the left edge, f the bottom, d the height.
        x = round(float(m.group(5)), 2)
        y = round(height - float(m.group(6)) - float(m.group(4)), 2)
        key = next((k for k in rows if abs(k - x) <= 6), x)
        rows.setdefault(key, []).append((x, y))
    for key, pts in rows.items():
        ys = sorted({y for _, y in pts})
        if len(ys) >= 3:
            out.append(((key, ys[0], key, ys[-1]), ys))
    return out


def pitch(ys: list[float]) -> float | None:
    steps = np.diff(ys)
    med = float(np.median(steps))
    keep = [y for y, s in zip(ys[1:], steps) if abs(s - med) < 1.0]
    if len(keep) < 2:
        return None
    keep = [ys[0]] + keep
    return float(np.polyfit(range(len(keep)), keep, 1)[0])


def main() -> None:
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    print(f"{'slide':>10} {'declared':>9} {'rendered':>9} {'lines':>5}  {'floor':>6} {'round':>6} {'half':>6}  text")
    seen: dict[tuple, int] = {}
    for arg in sys.argv[1:]:
        doc = f"{int(arg):02d}"
        src = ROOT / "pptx" / next(i["local"] for i in manifest if f"{i['idx']:02d}" == doc)
        pdf = pymupdf.open(ROOT / "ssim_pptx" / "ppt_pdf" / f"{doc}.pdf")
        with zipfile.ZipFile(src) as z:
            names = sorted((n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                           key=lambda n: int(re.findall(r"\d+", n)[-1]))
            for i, name in enumerate(names):
                if i >= len(pdf):
                    break
                blocks = pdf_blocks(pdf[i])
                for sh in shapes(z.read(name).decode("utf-8", "replace")):
                    best = None
                    for bbox, ys in blocks:
                        if (sh["x"] - 8 <= bbox[0] <= sh["x"] + sh["w"] + 8
                                and sh["y"] - 8 <= bbox[1] <= sh["y"] + sh["h"] + 8):
                            if best is None or len(ys) > len(best[1]):
                                best = (bbox, ys)
                    if best is None:
                        continue
                    p = pitch(best[1])
                    if p is None or p < 4:
                        continue
                    dec = sh["declared"]
                    mark = lambda v: "OK" if abs(v - p) < 0.12 else ""
                    print(f"{doc + '/s' + str(i + 1):>10} {dec:9.2f} {p:9.3f} {len(best[1]):5d}  "
                          f"{mark(int(dec)):>6} {mark(round(dec)):>6} {mark(round(dec * 2) / 2):>6}  {sh['text']!r}")
                    seen[(dec, round(p, 2))] = seen.get((dec, round(p, 2)), 0) + 1
    print("\ndeclared -> rendered, distinct pairs:")
    for (dec, p), n in sorted(seen.items()):
        print(f"  {dec:7.2f} -> {p:8.2f}   x{n}")


if __name__ == "__main__":
    main()
