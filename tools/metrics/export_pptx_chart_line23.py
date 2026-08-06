# -*- coding: utf-8 -*-
"""Export chart_line2 / chart_line3 to PDF via PowerPoint COM (fresh instance)."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

for name in ("chart_line2", "chart_line3"):
    base = os.path.abspath(os.path.join(r"pipeline_data\pptx_probes", name))
    pptx = os.path.join(base, name + ".pptx")
    pdf = os.path.join(base, name + ".pdf")
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(pptx, WithWindow=False)
        pres.SaveAs(pdf, 32)  # 32 = ppSaveAsPDF
        pres.Close()
        print("saved:", pdf, os.path.getsize(pdf))
    finally:
        app.Quit()
