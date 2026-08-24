# -*- coding: utf-8 -*-
"""Export the EOT-header variants to PDF with PowerPoint itself.

NEVER run this while the renderer is producing PNGs
(`pptx_com_render_must_not_overlap`).

Usage: python tools/metrics/export_pptx_italeot.py
"""
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_derive" / "italeot"

app = win32com.client.Dispatch("PowerPoint.Application")
try:
    for probe in sorted(OUT.glob("italeot_*.pptx")):
        pres = app.Presentations.Open(str(probe.resolve()), WithWindow=False)
        try:
            out = probe.with_suffix(".pdf")
            pres.SaveAs(str(out), 32)  # 32 = ppSaveAsPDF
            print("wrote", out)
        finally:
            pres.Close()
finally:
    app.Quit()
