# -*- coding: utf-8 -*-
"""Export the PATH (radial) gradient probe to PDF with PowerPoint itself.

NEVER run this while the renderer is producing PNGs -- a live PowerPoint COM
instance during a render run has corrupted whole decks before
(`pptx_com_render_must_not_overlap`).

Usage: python tools/metrics/export_pptx_gradpath.py
"""
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
PROBE = REPO / "pipeline_data" / "pptx_probes" / "gradpath" / "probe_gradpath.pptx"

app = win32com.client.Dispatch("PowerPoint.Application")
try:
    out = PROBE.with_suffix(".pdf")
    pres = app.Presentations.Open(str(PROBE.resolve()), WithWindow=False)
    try:
        pres.SaveAs(str(out), 32)  # 32 = ppSaveAsPDF
    finally:
        pres.Close()
    print("wrote", out)
finally:
    app.Quit()
