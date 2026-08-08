# -*- coding: utf-8 -*-
"""Export the doughnut-residual probe to PDF via PowerPoint COM."""
import sys, os

sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

base = os.path.abspath(r"pipeline_data\pptx_probes\chart_doughnut_resid")
src = os.path.join(base, "chart_doughnut_resid.pptx")
pdf = os.path.join(base, "chart_doughnut_resid.pdf")

app = win32com.client.DispatchEx("PowerPoint.Application")
try:
    pres = app.Presentations.Open(src, WithWindow=False)
    try:
        print("slides:", pres.Slides.Count)
        for i in range(1, pres.Slides.Count + 1):
            sl = pres.Slides(i)
            for sh in sl.Shapes:
                if sh.HasChart:
                    ch = sh.Chart
                    print(
                        f"  S{i}: type={ch.ChartType} series={ch.SeriesCollection().Count} "
                        f"title={ch.HasTitle} legend={ch.HasLegend}"
                    )
        if os.path.exists(pdf):
            os.remove(pdf)
        pres.SaveAs(pdf, 32)
        print("pdf:", pdf, os.path.getsize(pdf))
    finally:
        pres.Close()
finally:
    app.Quit()
