# -*- coding: utf-8 -*-
"""Export the line-end repro to PDF with PowerPoint itself.

NEVER run this while the renderer is producing PNGs: a live PowerPoint COM
instance during a render run has corrupted whole decks before (whole-deck
false regressions traced to nothing but the overlap).

Usage: python tools/metrics/export_pptx_lineend.py [deck.pptx]
"""
import sys
from pathlib import Path

import win32com.client

PROBE = Path(sys.argv[1] if len(sys.argv) > 1
             else r'pipeline_data\pptx_probe\probe_lineend.pptx').resolve()
out = PROBE.with_suffix('.pdf')

app = win32com.client.Dispatch('PowerPoint.Application')
pres = app.Presentations.Open(str(PROBE), WithWindow=False)
try:
    pres.SaveAs(str(out), 32)          # 32 = ppSaveAsPDF
finally:
    pres.Close()
    app.Quit()
print('wrote', out)
