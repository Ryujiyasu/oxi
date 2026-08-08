# -*- coding: utf-8 -*-
"""Convert chart_bar.pptx to PDF via PowerPoint COM and dump chart COM props."""
import sys, os

sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

pptx = os.path.abspath(r"pipeline_data\pptx_probes\chart_bar\chart_bar.pptx")
pdf = os.path.abspath(r"pipeline_data\pptx_probes\chart_bar\chart_bar.pdf")

ppt = win32com.client.DispatchEx("PowerPoint.Application")
try:
    pres = ppt.Presentations.Open(pptx, WithWindow=False)
    pres.SaveAs(pdf, 32)  # ppSaveAsPDF
    for i in range(1, pres.Slides.Count + 1):
        sh = pres.Slides(i).Shapes(1)
        info = {"slide": i}
        if sh.HasChart:
            ch = sh.Chart
            info["chart_type"] = int(ch.ChartType)
            info["has_title"] = bool(ch.HasTitle)
            info["has_legend"] = bool(ch.HasLegend)
            sc = ch.SeriesCollection()
            info["series_count"] = sc.Count
            for k in range(1, sc.Count + 1):
                info[f"s{k}"] = sc(k).Name
        print(info)
    pres.Close()
finally:
    ppt.Quit()
print("saved:", pdf, os.path.getsize(pdf))
