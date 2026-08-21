# -*- coding: utf-8 -*-
"""Derive each `a:prstDash` preset's dash pattern from PowerPoint's own PDF.

A PDF stroke carries its dash array verbatim (`[on off] phase d`), and
PowerPoint writes the pattern it drew — so the preset's ratios can be read
rather than guessed. This pairs every non-solid line in the dev decks with the
PDF strokes on its slide and reports the array in units of the line width,
which is how DrawingML defines the presets.

Usage: python tools/metrics/read_pptx_dash.py [--decks d06,d34]
"""
from __future__ import annotations

import argparse
import glob
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev"


def slide_dashes(pptx: Path) -> dict[int, list[tuple[str, float]]]:
    """slide number -> [(preset, line width pt)] for every non-solid line."""
    out: dict[int, list[tuple[str, float]]] = defaultdict(list)
    z = zipfile.ZipFile(pptx)
    for name in z.namelist():
        m = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
        if not m:
            continue
        xml = z.read(name).decode("utf-8", "replace")
        # Pair each a:ln with the prstDash inside it.
        for ln in re.finditer(r"<a:ln\b([^>]*)>(.*?)</a:ln>", xml, re.S):
            body = ln.group(2)
            dash = re.search(r'<a:prstDash val="([^"]+)"', body)
            if not dash or dash.group(1) == "solid":
                continue
            w = re.search(r'w="(\d+)"', ln.group(1))
            out[int(m.group(1))].append(
                (dash.group(1), (int(w.group(1)) / 12700.0) if w else 0.75)
            )
    return out


def pdf_dashes(pdf_path: Path, page_no: int) -> list[tuple[str, float]]:
    """[(dash array string, stroke width pt)] for every dashed stroke."""
    pdf = pymupdf.open(pdf_path)
    try:
        page = pdf[page_no - 1]
        seen = []
        for d in page.get_drawings():
            dashes = d.get("dashes") or ""
            if not dashes or dashes.strip() in ("[] 0", "[]"):
                continue
            seen.append((dashes.strip(), float(d.get("width") or 0.0)))
        return seen
    finally:
        pdf.close()


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--decks", default=None)
    args = ap.parse_args()
    wanted = {s.strip() for s in args.decks.split(",")} if args.decks else None

    per_preset: dict[str, list[tuple[str, float, float]]] = defaultdict(list)
    for pptx in sorted((DEV / "pptx").glob("*.pptx")):
        did = pptx.stem.split("__")[0]
        if wanted and did not in wanted:
            continue
        pdf_hits = glob.glob(str(DEV / "pdf" / f"{pptx.stem}.pdf"))
        if not pdf_hits:
            continue
        declared = slide_dashes(pptx)
        for slide_no, entries in sorted(declared.items()):
            drawn = pdf_dashes(Path(pdf_hits[0]), slide_no)
            presets = {p for p, _ in entries}
            if len(presets) == 1 and drawn:
                preset = next(iter(presets))
                for arr, w in drawn:
                    per_preset[preset].append((arr, w, entries[0][1]))
            else:
                print(f"  {did} s{slide_no}: declared {sorted(presets)} "
                      f"({len(entries)} lines) / drawn {len(drawn)} dashed strokes"
                      f"{' AMBIGUOUS' if len(presets) > 1 else ''}")
    print()
    for preset, rows in sorted(per_preset.items()):
        print(f"== {preset} ({len(rows)} strokes)")
        tally: dict[str, int] = defaultdict(int)
        for arr, w, declared_w in rows:
            nums = [float(x) for x in re.findall(r"[-\d.]+", arr)]
            ratio = " ".join(f"{n / w:.3g}" for n in nums[:-1]) if w else "?"
            tally[f"{arr}  w={w:.3g}pt  = {ratio} x width (xml w={declared_w:.3g})"] += 1
        for k, n in sorted(tally.items(), key=lambda kv: -kv[1])[:6]:
            print(f"   x{n}  {k}")


if __name__ == "__main__":
    main()
