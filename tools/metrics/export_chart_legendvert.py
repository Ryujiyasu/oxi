# -*- coding: utf-8 -*-
"""Export the legend-vertical probe to PDF via PowerPoint COM."""
import sys, os

sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

base = os.path.abspath(r"pipeline_data\pptx_probes\chart_legendvert")
pptx = os.path.join(base, "chart_legendvert.pptx")
pdf = os.path.join(base, "chart_legendvert.pdf")

app = win32com.client.DispatchEx("PowerPoint.Application")
try:
    pres = app.Presentations.Open(pptx, WithWindow=False)
    print("slides:", pres.Slides.Count)
    pres.SaveAs(pdf, 32)
    pres.Close()
finally:
    app.Quit()
print("pdf:", pdf, os.path.getsize(pdf))
