# -*- coding: utf-8 -*-
"""Which face did PowerPoint draw each arm of bzero.pptx with?"""
import os
import sys

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ARMS = [
    "level-bold, run b=0 beside a silent run",
    "level-bold, run b=1 beside a silent run",
    "level not bold, run b=0 beside a silent run",
    "level-bold, both runs in ONE paragraph",
]
pdf = os.path.join("pipeline_data", "pptx_probes", "bzero", "bzero.pdf")
doc = pymupdf.open(pdf)
for i in range(len(doc)):
    print(f"slide {i+1}: {ARMS[i] if i < len(ARMS) else ''}")
    for block in doc[i].get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for s in line["spans"]:
                print(f"   {s['font']:18} {s['size']:6.2f}pt  x={s['bbox'][0]:7.2f}  "
                      f"{s['text']!r}")
