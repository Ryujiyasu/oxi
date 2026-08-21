# -*- coding: utf-8 -*-
"""Read the embedbold probe: which face does PowerPoint draw, and where does a
cell's baseline land, for an EMBEDDED weight-named family?

Prints, per slide:
  * what the deck embedded for that typeface
  * the PDF font each span uses, for the b=0 and b=1 free-text lines
  * the cell first-baseline offset A, with the row grown to the confirmed
    2*margin + 1.2*size so the block top is row_top + margin

Usage: python tools/metrics/read_pptx_embedbold.py [subdir]
"""
from __future__ import annotations

import io
import json
import re
import sys
import zipfile
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DIR = Path(r"pipeline_data\pptx_probes\embedbold").resolve()


def embedded_styles(pptx: Path) -> dict[str, list[str]]:
    z = zipfile.ZipFile(pptx)
    pres = z.read("ppt/presentation.xml").decode("utf-8", "replace")
    out: dict[str, list[str]] = {}
    for m in re.finditer(r"<p:embeddedFont>.*?</p:embeddedFont>", pres, re.S):
        blk = m.group(0)
        face = re.search(r'typeface="([^"]+)"', blk)
        if not face:
            continue
        out[face.group(1)] = re.findall(r"<p:(regular|bold|italic|boldItalic)\b", blk)
    return out


def subset_metrics(pdf: pymupdf.Document, xref: int) -> tuple[str, float] | None:
    """(name, 1.2*asc/(asc+desc)) for one embedded PDF subset."""
    try:
        from fontTools.ttLib import TTFont
    except ImportError:
        return None
    name, _, _, buf = pdf.extract_font(xref)
    if not buf:
        return None
    try:
        f = TTFont(io.BytesIO(buf), lazy=True, fontNumber=0)
        upm = f["head"].unitsPerEm
        os2 = f["OS/2"]
    except Exception:
        return None
    if os2.fsSelection & 0x80:
        asc = (os2.sTypoAscender + os2.sTypoLineGap) / upm
        desc = -os2.sTypoDescender / upm
    else:
        asc = os2.usWinAscent / upm
        desc = os2.usWinDescent / upm
    return name, 1.2 * asc / (asc + desc)


def main() -> None:
    sub = sys.argv[1] if len(sys.argv) > 1 else "emb"
    manifest = json.loads((DIR / "embedbold_manifest.json").read_text(encoding="utf-8"))
    embedded = embedded_styles(DIR / "embedbold_embedded.pptx")
    print("deck embeds:", {k: v for k, v in embedded.items()})
    pdf = pymupdf.open(DIR / sub / "deck.pdf")
    for row in manifest:
        page = pdf[row["slide"] - 1]
        print(f"\n== {row['face']} @ {row['size']}pt")
        models = {}
        for xref, *_rest in page.get_fonts():
            m = subset_metrics(pdf, xref)
            if m:
                models[m[0]] = m[1]
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for line in b["lines"]:
                for s in line["spans"]:
                    ch = s["chars"]
                    txt = "".join(c["c"] for c in ch).strip()
                    if not txt or txt.startswith(row["face"]):
                        continue
                    y = ch[0]["origin"][1]
                    note = ""
                    if txt.startswith("Hxy"):
                        # the table's rows are grown to 2*margin + 1.2*size, so
                        # the block top of row i is table_top + i*grown + margin
                        grown = 2 * row["margin"] + 1.2 * row["size"]
                        i = round((y - row["table_top"] - row["margin"]
                                   - 0.97 * row["size"]) / grown)
                        top = row["table_top"] + i * grown + row["margin"]
                        note = f"  cell row {i}: A = {(y - top) / row['size']:.4f}"
                    print(f"   {s['font']:26s} y={y:7.2f} {txt[:26]!r}{note}")
        print("   subsets ->", {k: round(v, 4) for k, v in models.items()})


if __name__ == "__main__":
    main()
