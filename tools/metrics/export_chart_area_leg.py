# -*- coding: utf-8 -*-
"""Export chart_area_leg.pptx to PDF via PowerPoint COM."""
import sys, os

sys.stdout.reconfigure(encoding="utf-8")
import win32com.client as win32

base = os.path.abspath(r"pipeline_data\pptx_probes\chart_area_leg")
src = os.path.join(base, "chart_area_leg.pptx")
pdf = os.path.join(base, "chart_area_leg.pdf")

app = win32.DispatchEx("PowerPoint.Application")
try:
    pres = app.Presentations.Open(src, WithWindow=False)
    for i in range(1, pres.Slides.Count + 1):
        sl = pres.Slides(i)
        for sh in sl.Shapes:
            if sh.HasChart:
                c = sh.Chart
                print(
                    f"S{i}: type={c.ChartType} series={c.SeriesCollection().Count} "
                    f"title={c.HasTitle} legend={c.HasLegend}"
                )
    if os.path.exists(pdf):
        os.remove(pdf)
    pres.SaveAs(pdf, 32)
    pres.Close()
finally:
    app.Quit()
print("saved:", pdf, os.path.getsize(pdf))
