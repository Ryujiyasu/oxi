# -*- coding: utf-8 -*-
"""Read the slide-number probe PDFs: what did PowerPoint print in each field?

Usage: python tools/metrics/read_pptx_slidenum.py
"""
import re
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
DIR = REPO / "pipeline_data" / "pptx_probes" / "slidenum"
ARMS = {"a0": None, "a1": 5, "a2": 100}

for tag, first in ARMS.items():
    pdf = DIR / f"probe_slidenum_{tag}.pdf"
    if not pdf.exists():
        print(f"{tag}: no pdf"); continue
    doc = pymupdf.open(pdf)
    print(f"-- {tag}  firstSlideNum={first}")
    base = first if first is not None else 1
    for i in range(len(doc)):
        txt = doc[i].get_text().replace("\n", " ")
        m = re.search(r"\[\s*(.*?)\s*\]", txt)
        printed = m.group(1) if m else txt.strip()
        want = str(i + base)
        flag = "ok" if printed == want else "MISMATCH"
        cache = " (stale cache 777)" if i == 2 else ""
        print(f"   slide {i+1}: printed '{printed}'  model '{want}'  {flag}{cache}")
