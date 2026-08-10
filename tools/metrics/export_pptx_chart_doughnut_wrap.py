# -*- coding: utf-8 -*-
"""Export the wrap probe to PDF via PowerPoint COM (fresh instance).

Saves to chart_doughnut_wrap_ppt.pdf so the existing Word/LibreOffice
reference PDF (chart_doughnut_wrap.pdf) is not overwritten.
"""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

base = os.path.abspath(r"pipeline_data\pptx_probes\chart_doughnut_wrap")
pptx = os.path.join(base, "chart_doughnut_wrap.pptx")
pdf = os.path.join(base, "chart_doughnut_wrap_ppt.pdf")

app = win32com.client.DispatchEx("PowerPoint.Application")
try:
    pres = app.Presentations.Open(pptx, WithWindow=False)
    pres.SaveAs(pdf, 32)  # ppSaveAsPDF
    print("pages:", len(pres.Slides))
    pres.Close()
finally:
    app.Quit()
print("saved:", pdf, os.path.getsize(pdf))
