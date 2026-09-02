# -*- coding: utf-8 -*-
"""Where does a deck's truth PDF draw Calibri that the deck never asked for?

`gen_pptx_missfont.py` put the question to PowerPoint directly -- a deck themed
Georgia draws an unservable family in **Calibri**, not in its theme face -- and
this is the corpus half of that claim: every deck whose theme is NOT Calibri,
whose slides never name Calibri, and whose own PDF contains Calibri anyway, is
a deck where PowerPoint substituted it.

It also lists what those decks DO name, because the substituted family is the
one Oxi has to stop drawing: `pptx_cff_part_census.py` found PowerPoint refuses
a CFF-outlined embedded part, and the naive fix (drop the part, let GDI choose)
cost blind 31 -0.0199 SSIM because GDI's substitute is not Calibri either.

    python tools/metrics/pptx_fallback_census.py
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"


def theme_faces(z: zipfile.ZipFile) -> list[str]:
    out: list[str] = []
    for name in z.namelist():
        if name.startswith("ppt/theme/"):
            text = z.read(name).decode("utf-8", "replace")
            scheme = re.search(r"<a:fontScheme.*?</a:fontScheme>", text, re.S)
            if scheme:
                out += re.findall(r'<a:latin typeface="([^"]*)"', scheme.group(0))
    return sorted({f for f in out if f})


def named_faces(z: zipfile.ZipFile) -> set[str]:
    """Every typeface the slides, layouts and masters name outright."""
    out: set[str] = set()
    for name in z.namelist():
        if re.match(r"ppt/(slides|slideLayouts|slideMasters)/[^/]+\.xml$", name):
            text = z.read(name).decode("utf-8", "replace")
            out |= set(re.findall(r'typeface="([^"]*)"', text))
    return {f for f in out if f and not f.startswith("+")}


def main() -> None:
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    hits = read = themed_calibri = names_calibri = eligible = 0
    for item in manifest:
        doc = f"{item['idx']:02d}"
        src = ROOT / "pptx" / item["local"]
        pdf_path = ROOT / "ssim_pptx" / "ppt_pdf" / f"{doc}.pdf"
        if not src.exists() or not pdf_path.exists():
            continue
        try:
            with zipfile.ZipFile(src) as z:
                theme = theme_faces(z)
                named = named_faces(z)
        except Exception as e:                                   # a corrupt deck
            print(f"{doc}: unreadable ({str(e)[:40]})", flush=True)
            continue
        read += 1
        if any(t.lower().startswith("calibri") for t in theme):
            themed_calibri += 1
            continue                                             # cannot tell them apart
        if any(n.lower().startswith("calibri") for n in named):
            names_calibri += 1
            continue                                             # asked for it outright
        eligible += 1
        doc_pdf = pymupdf.open(pdf_path)
        pages = [
            p + 1
            for p in range(doc_pdf.page_count)
            if any(f[3].split("+")[-1].lower().startswith("calibri")
                   for f in doc_pdf[p].get_fonts())
        ]
        doc_pdf.close()
        if pages:
            hits += 1
            print(f"{doc}: theme {theme}, never names Calibri, yet its PDF draws it on "
                  f"{len(pages)} page(s): {pages[:8]}", flush=True)
            print(f"      names: {sorted(named)[:10]}", flush=True)
    # ★The denominator, always -- a bare "0 hits" from a test nothing was
    # ELIGIBLE for reads like evidence and is not any. This corpus turns out to
    # answer nothing here: 45 of its 48 decks name Calibri in some layout or
    # master, so Calibri in the PDF is never a substitution one can attribute.
    print(f"\n{read} decks read: {themed_calibri} themed Calibri, "
          f"{names_calibri} name it outright, {eligible} could be tested")
    print(f"{hits} of those {eligible} draw Calibri without asking for it")
    if not eligible:
        print("This corpus cannot answer the question -- see the probe "
              "(`gen_pptx_missfont.py`), which asks PowerPoint directly.")


if __name__ == "__main__":
    main()
