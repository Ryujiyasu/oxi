# -*- coding: utf-8 -*-
"""The text rectangle each preset gives, solved from the three alignments.

    algn="l"    pen x  =  text_left
    algn="r"    pen x  =  text_right - line_w
    algn="ctr"  pen x  =  text_left + (text_width - line_w) / 2   (a check)

Reported as a FRACTION of the bounding box, so the numbers can be compared
against the published preset formulas and used for any box size.
"""
import os
import sys

import pymupdf
from fontTools.ttLib import TTFont

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TEXT = "Wq"
SIZE_PT = 18.0
BOX_L, BOX_W = 200.0, 300.0
PRESETS = [
    ("rect", []), ("ellipse", []), ("homePlate", []),
    ("homePlate", [("adj", 30129)]), ("homePlate", [("adj", 50000)]),
    ("teardrop", []), ("pie", []), ("roundRect", []),
    ("chevron", []), ("wedgeRectCallout", []),
]
ALIGNS = ["l", "r", "ctr"]

t = TTFont(r"C:\Windows\Fonts\arial.ttf", lazy=True)
upm, hm, cm = t["head"].unitsPerEm, t["hmtx"], t.getBestCmap()
line_w = sum(hm[cm[ord(c)]][0] for c in TEXT) / upm * SIZE_PT

pdf = os.path.join("pipeline_data", "pptx_probes", "textrect", "textrect.pdf")
if not os.path.exists(pdf):
    sys.exit(f"no {pdf} -- run export_pptx_textrect.py first")
doc = pymupdf.open(pdf)


def pen_x(page):
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for s in line["spans"]:
                if s["text"].strip() == TEXT:
                    return s["origin"][0]
    return None


print(f"{TEXT!r} advances {line_w:.3f}pt at {SIZE_PT:g}pt; "
      f"box left {BOX_L:g}pt width {BOX_W:g}pt\n")
print(f"{'preset':24}{'left':>9}{'right':>9}{'width':>9}"
      f"{'l/w':>8}{'r/w':>8}   centre check")
i = 0
for preset, adjs in PRESETS:
    got = {}
    for algn in ALIGNS:
        got[algn] = pen_x(doc[i])
        i += 1
    if any(v is None for v in got.values()):
        print(f"{preset:24}   not found")
        continue
    left = got["l"] - BOX_L
    right = got["r"] - BOX_L + line_w
    width = right - left
    ctr_pred = BOX_L + left + (width - line_w) / 2
    name = preset + (f" adj={adjs[0][1]}" if adjs else "")
    print(f"{name:24}{left:9.3f}{right:9.3f}{width:9.3f}"
          f"{left/BOX_W:8.4f}{right/BOX_W:8.4f}   "
          f"{got['ctr'] - ctr_pred:+.3f}pt")
