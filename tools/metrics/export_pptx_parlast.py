# -*- coding: utf-8 -*-
"""Export the parlast probe to PDF with PowerPoint itself.

NEVER run this while the renderer is producing PNGs (pptx_com_render_must_not_overlap).

    python tools/metrics/export_pptx_parlast.py
"""
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
PROBE = REPO / "pipeline_data" / "pptx_probes" / "parlast" / "probe_parlast.pptx"

app = win32com.client.Dispatch("PowerPoint.Application")
try:
    out = PROBE.with_suffix(".pdf")
    pres = app.Presentations.Open(str(PROBE.resolve()), WithWindow=False)
    try:
        pres.SaveAs(str(out), 32)  # ppSaveAsPDF
    finally:
        pres.Close()
    print("wrote", out)
finally:
    app.Quit()
