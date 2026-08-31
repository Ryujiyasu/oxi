# -*- coding: utf-8 -*-
"""Where PowerPoint drew each arm of overwide.pptx, against both predictions."""
import os
import sys

import pymupdf
from fontTools.ttLib import TTFont

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BOX_LEFT = 300.0
SIZE_PT = 100.0
ARMS = [("ctr", 40.0), ("r", 40.0), ("l", 40.0), ("ctr", 300.0), ("r", 300.0)]

t = TTFont(r"C:\Windows\Fonts\arial.ttf", lazy=True)
adv = t["hmtx"][t.getBestCmap()[ord("W")]][0] / t["head"].unitsPerEm * SIZE_PT
print(f"'W' advances {adv:.3f}pt at {SIZE_PT:g}pt")

pdf = os.path.join("pipeline_data", "pptx_probes", "overwide", "overwide.pdf")
if not os.path.exists(pdf):
    sys.exit(f"no {pdf} -- run export_pptx_overwide.py first")
doc = pymupdf.open(pdf)
print(f"{'arm':>4} {'algn':>5} {'box_w':>7} {'truth':>10} {'unclamped':>11}"
      f" {'clamped':>9}   verdict")
for i, (algn, box_w) in enumerate(ARMS):
    got = None
    for block in doc[i].get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for s in line["spans"]:
                if s["text"].strip() == "W":
                    got = s["origin"][0]
    if got is None:
        print(f"{i+1:4} {algn:>5} {box_w:7g}   not found")
        continue
    off = (box_w - adv) / 2 if algn == "ctr" else (box_w - adv) if algn == "r" else 0.0
    unclamped = BOX_LEFT + off
    clamped = BOX_LEFT + max(off, 0.0)
    which = ("unclamped" if abs(got - unclamped) < abs(got - clamped)
             else "clamped" if abs(got - clamped) < abs(got - unclamped) else "same")
    print(f"{i+1:4} {algn:>5} {box_w:7g} {got:10.3f} {unclamped:11.3f}"
          f" {clamped:9.3f}   {which} ({min(abs(got-unclamped), abs(got-clamped)):+.3f})")
