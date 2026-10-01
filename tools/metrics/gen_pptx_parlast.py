# -*- coding: utf-8 -*-
"""Probe: is a paragraph's LAST line placed by a different ascent split?

d32 / d39 / d44 set titles in embedded Bebas Neue, whose usWin split is 0.8769
and whose hhea/typo split is 0.9000. PowerPoint puts every line of those
paragraphs where the shipped first-baseline rule computes with usWin -- except
the paragraph's last line, which sits 0.0235 em lower (= 0.9000 - 0.8769). A
one-line paragraph's only line IS its last line, which is why one-line titles
looked "hhea" and multi-line ones "win" in the corpus scorer.

Each arm is one paragraph of 1-3 lines (hard `a:br` breaks or a narrow box that
wraps) in a fixed box: `anchor=t`, insets 0, `noAutofit`, so every baseline can
be read from PowerPoint's PDF as an offset from the box top. Fonts are installed
faces whose hhea, typo and usWin splits all DIFFER (Palatino, Candara,
Consolas), whose hhea exceeds usWin the way Bebas Neue's does (Magneto, Old
English Text MT), and Arial as the control where the three agree.

    python tools/metrics/gen_pptx_parlast.py
    python tools/metrics/export_pptx_parlast.py   # PowerPoint COM -> PDF
    python tools/metrics/read_pptx_parlast.py     # read baselines back
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_probes" / "parlast"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

BOX_X = Emu(457200)      # 36pt
BOX_Y = Emu(914400)      # 72pt -- the reader's origin
BOX_H = Emu(5486400)     # 432pt, room for three 60pt lines at 120%
SIZE = 60
FONTS = ["Arial", "Palatino Linotype", "Candara", "Consolas", "Magneto", "Old English Text MT"]
SPACINGS = [80, 100, 120]
WORDS = ["HAMBURG", "MINDEN", "BAXTER"]

ARMS = []
for font in FONTS:
    for pct in SPACINGS:
        for lines in (1, 2, 3):
            ARMS.append({"font": font, "pct": pct, "lines": lines, "mode": "br"})
        # the same two lines by wrapping instead of a hard break
        ARMS.append({"font": font, "pct": pct, "lines": 2, "mode": "wrap"})


def body_xml(arm: dict) -> str:
    rpr = (f'<a:rPr lang="en-US" sz="{SIZE * 100}" dirty="0">'
           f'<a:latin typeface="{arm["font"]}"/></a:rPr>')
    words = WORDS[:arm["lines"]]
    if arm["mode"] == "br":
        runs = f'<a:br>{rpr}</a:br>'.join(f'<a:r>{rpr}<a:t>{w}</a:t></a:r>' for w in words)
    else:
        runs = f'<a:r>{rpr}<a:t>{" ".join(words)}</a:t></a:r>'
    return (
        f'<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:a="{A}">'
        '<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></a:bodyPr>'
        '<a:lstStyle/>'
        f'<a:p><a:pPr><a:lnSpc><a:spcPct val="{arm["pct"] * 1000}"/></a:lnSpc>'
        '<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr>'
        f'{runs}<a:endParaRPr lang="en-US" sz="{SIZE * 100}"><a:latin typeface="{arm["font"]}"/></a:endParaRPr></a:p>'
        '</p:txBody>'
    )


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    blank = prs.slide_layouts[6]
    for arm in ARMS:
        slide = prs.slides.add_slide(blank)
        # a wrap arm gets a box narrower than two words, wider than either one
        width = Emu(11277600) if arm["mode"] == "br" else Emu(4114800)
        box = slide.shapes.add_textbox(BOX_X, BOX_Y, width, BOX_H)
        sp = box._element
        old = sp.find("{http://schemas.openxmlformats.org/presentationml/2006/main}txBody")
        sp.replace(old, etree.fromstring(body_xml(arm)))
    path = OUT / "probe_parlast.pptx"
    prs.save(path)
    (OUT / "arms.json").write_text(json.dumps(
        {"box_top_pt": BOX_Y / 12700, "size": SIZE, "arms": ARMS}, indent=1), encoding="utf-8")
    print("wrote", path, len(ARMS), "arms")


if __name__ == "__main__":
    main()
