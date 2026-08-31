# -*- coding: utf-8 -*-
"""Export textrect.pptx -> textrect.pdf via PowerPoint COM."""
import os
import sys

import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

name = "textrect"
base = os.path.abspath(os.path.join("pipeline_data", "pptx_probes", name))
pptx = os.path.join(base, name + ".pptx")
pdf = os.path.join(base, name + ".pdf")

app = win32com.client.DispatchEx("PowerPoint.Application")
try:
    pres = app.Presentations.Open(pptx, WithWindow=False)
    pres.SaveAs(pdf, 32)  # ppSaveAsPDF
    pres.Close()
finally:
    app.Quit()
print("saved:", pdf, os.path.getsize(pdf))
