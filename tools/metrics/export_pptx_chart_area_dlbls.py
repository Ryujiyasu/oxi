# -*- coding: utf-8 -*-
"""Export the AREA data-label probe to PDF via PowerPoint COM (fresh
instance) and record the COM truth for each slide's chart."""
import sys, os, json

sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

base = os.path.abspath(r"pipeline_data\pptx_probes\chart_area_dlbls")
pptx = os.path.join(base, "chart_area_dlbls.pptx")
pdf = os.path.join(base, "chart_area_dlbls.pdf")
truth = os.path.join(base, "chart_area_dlbls_truth.json")

app = win32com.client.DispatchEx("PowerPoint.Application")
rows = []
try:
    pres = app.Presentations.Open(pptx, WithWindow=False)
    print("slide size:", pres.PageSetup.SlideWidth, "x", pres.PageSetup.SlideHeight)
    for i in range(1, len(pres.Slides) + 1):
        sl = pres.Slides(i)
        for sh in sl.Shapes:
            if not sh.HasChart:
                continue
            ch = sh.Chart
            r = {
                "slide": i,
                "left": float(sh.Left), "top": float(sh.Top),
                "width": float(sh.Width), "height": float(sh.Height),
                "chart_type": int(ch.ChartType),
                "has_title": bool(ch.HasTitle),
                "has_legend": bool(ch.HasLegend),
                "series_count": int(ch.SeriesCollection().Count),
            }
            rows.append(r)
            print(r)
    pres.SaveAs(pdf, 32)  # ppSaveAsPDF
    pres.Close()
finally:
    app.Quit()

with open(truth, "w", encoding="utf-8") as f:
    json.dump(rows, f, ensure_ascii=False, indent=1)
print("saved:", pdf, os.path.getsize(pdf))
print("truth:", truth)
