# -*- coding: utf-8 -*-
"""Export the line-chart probe to PDF via PowerPoint COM (fresh instance)."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

base = os.path.abspath(r"pipeline_data\pptx_probes\chart_line_probe")
pptx = os.path.join(base, "chart_line_probe.pptx")
pdf = os.path.join(base, "chart_line_probe.pdf")

app = win32com.client.DispatchEx("PowerPoint.Application")
try:
    pres = app.Presentations.Open(pptx, WithWindow=False)
    pres.SaveAs(pdf, 32)  # 32 = ppSaveAsPDF
    pres.Close()
    print("saved:", pdf, os.path.getsize(pdf))
finally:
    app.Quit()
