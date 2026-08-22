# -*- coding: utf-8 -*-
"""Export the slide-number probe decks to PDF with PowerPoint itself.

NEVER run this while the renderer is producing PNGs -- a live PowerPoint COM
instance during a render run has corrupted whole decks before.

Usage: python tools/metrics/export_pptx_slidenum.py
"""
import sys
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
PROBES = sorted((REPO / "pipeline_data" / "pptx_probes" / "slidenum").glob("probe_slidenum_a*.pptx"))

app = win32com.client.Dispatch("PowerPoint.Application")
try:
    for probe in PROBES:
        out = probe.with_suffix(".pdf")
        pres = app.Presentations.Open(str(probe.resolve()), WithWindow=False)
        try:
            pres.SaveAs(str(out), 32)  # 32 = ppSaveAsPDF
        finally:
            pres.Close()
        print("wrote", out)
finally:
    app.Quit()
