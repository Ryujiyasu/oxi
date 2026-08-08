# -*- coding: utf-8 -*-
"""Export the area-chart probe to PDF and dump COM chart facts."""
import sys, os

sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

base = os.path.abspath(r"pipeline_data\pptx_probes\chart_area")
pptx = os.path.join(base, "chart_area.pptx")
pdf = os.path.join(base, "chart_area.pdf")

app = win32com.client.DispatchEx("PowerPoint.Application")
try:
    pres = app.Presentations.Open(pptx, WithWindow=False)
    for i in range(1, pres.Slides.Count + 1):
        sl = pres.Slides(i)
        for sh in sl.Shapes:
            if sh.HasChart:
                c = sh.Chart
                print(
                    f"S{i}: type={c.ChartType} series={c.SeriesCollection().Count} "
                    f"title={c.HasTitle} legend={c.HasLegend}"
                )
    pres.SaveAs(pdf, 32)
    pres.Close()
finally:
    app.Quit()
print("pdf:", pdf, os.path.getsize(pdf))
