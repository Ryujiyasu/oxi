# -*- coding: utf-8 -*-
"""The advance PowerPoint itself uses, asked of PowerPoint and not of a PDF.

`read_pptx_drawgrid.py` asks the truth PDF and gets a mess: PowerPoint restates
the geometry there as `Tf` (a size that is NOT the declared one -- 8.04 for 8pt,
18.024 for 18pt), plus a per-run `Tc`, plus sparse integer `TJ` corrections, and
the effective advance for one glyph in one face wobbles +-0.9% with the declared
size in a way no grid, scale or hinted device resolution (96..2400 dpi) fits.
**The truth PDF is not a per-glyph advance oracle.**

PowerPoint's own measurement API is. `BoundWidth` of a run of N identical glyphs
is N*advance plus one glyph's ink overhang, which does not depend on N, so
(BW(N2) - BW(N1)) / (N2 - N1) is the advance alone.

★THE ANSWER (2026-09-01), 3 faces x 8 sizes:

    advance == round(em * size * 8) / 8        22 of 24 arms, exactly
    (Verdana 32pt is a 23rd: both models agree there and the tie reads as
     "neither"; Verdana 8pt is the one real exception, +1 master unit above
     the design advance, which is what a strongly hinted face does at a small
     ppem.)

So the master unit (1/8pt, 576 to the inch) is not only what PowerPoint BREAKS
on (`pptx-master-unit-break-law`) -- it is what PowerPoint MEASURES and DRAWS
on. Every one of the 24 `BoundWidth` values is itself an exact multiple of
1/8pt, including at fractional point sizes (12.5 -> 7.00000, 15.99 -> 8.87500).

Usage: python tools/metrics/read_pptx_drawgrid_com.py
NEVER while the renderer is producing PNGs (`pptx_com_render_must_not_overlap`).
"""
import json, os, sys
from pathlib import Path
import win32com.client
from fontTools.ttLib import TTFont
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_probes" / "drawgrid"
FILES = {"Arial": "arial.ttf", "Georgia": "georgia.ttf", "Verdana": "verdana.ttf"}

def em_of(face, ch):
    f = TTFont(os.path.join(os.environ["WINDIR"], "Fonts", FILES[face]),
               lazy=True, checkChecksums=0)
    return f["hmtx"][f.getBestCmap()[ord(ch)]][0] / f["head"].unitsPerEm

arms = json.loads((OUT / "arms.json").read_text(encoding="utf-8"))
COUNTS = [20, 60]
app = win32com.client.Dispatch("PowerPoint.Application")
try:
    pres = app.Presentations.Open(str((OUT / "probe_drawgrid.pptx").resolve()),
                                  WithWindow=False)
    print(f"{'face':9}{'sz':>7}{'BW20':>10}{'BW60':>10}{'advance':>10}"
          f"{'exact':>10}{'grid1/8':>10}  verdict")
    tally = {"GRID": 0, "exact": 0, "neither": 0}
    try:
        for a in arms:
            tr = pres.Slides(a["slide"]).Shapes(1).TextFrame.TextRange
            bw = []
            for n in COUNTS:
                tr.Text = "n" * n
                bw.append(tr.BoundWidth)
            adv = (bw[1] - bw[0]) / (COUNTS[1] - COUNTS[0])
            em = em_of(a["typeface"], "n")
            exact = em * a["sz_pt"]
            grid = round(em * a["sz_pt"] * 8) / 8
            de, dg = abs(adv - exact), abs(adv - grid)
            v = ("GRID" if dg < de and dg < 0.002 else
                 "exact" if de < dg and de < 0.002 else "neither")
            tally[v] += 1
            print(f"{a['typeface']:9}{a['sz_pt']:7.2f}{bw[0]:10.4f}{bw[1]:10.4f}"
                  f"{adv:10.5f}{exact:10.5f}{grid:10.5f}  {v}")
    finally:
        pres.Saved = True
        pres.Close()
    print(f"\n{tally}")
finally:
    app.Quit()
