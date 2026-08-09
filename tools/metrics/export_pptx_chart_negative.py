# -*- coding: utf-8 -*-
"""Export the negative-value probe to PDF via PowerPoint COM."""
import os
import win32com.client

BASE = os.path.abspath(os.path.join("pipeline_data", "pptx_probes", "chart_negative"))
PPTX = os.path.join(BASE, "chart_negative.pptx")
PDF = os.path.join(BASE, "chart_negative.pdf")

app = win32com.client.DispatchEx("PowerPoint.Application")
try:
    pres = app.Presentations.Open(PPTX, WithWindow=False)
    try:
        print("slides:", pres.Slides.Count)
        for i in range(1, pres.Slides.Count + 1):
            sl = pres.Slides(i)
            for sh in sl.Shapes:
                if sh.HasChart:
                    ch = sh.Chart
                    print("  S%d chart_type=%s has_title=%s has_legend=%s series=%d"
                          % (i, ch.ChartType, ch.HasTitle, ch.HasLegend,
                             ch.SeriesCollection().Count))
        pres.SaveAs(PDF, 32)
    finally:
        pres.Close()
finally:
    app.Quit()
print("saved", PDF)
