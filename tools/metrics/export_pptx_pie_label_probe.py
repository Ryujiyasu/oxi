# -*- coding: utf-8 -*-
"""Export pie_label_probe.pptx -> pie_label_probe.pdf via PowerPoint COM."""
import os, sys
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

name = "pie_label_probe"
base = os.path.abspath(os.path.join(r"pipeline_data\pptx_probes", name))
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
