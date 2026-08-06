# -*- coding: utf-8 -*-
"""Word COM truth for the line chart probe: shape geometry + chart type +
auto title / legend state (PowerPoint COM), then export the PDF (absolute
path) for fitz measurement."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import json
import win32com.client

base = os.path.abspath(r"pipeline_data\pptx_probes\chart_line")
pptx = os.path.join(base, "chart_line.pptx")
pdf = os.path.join(base, "chart_line.pdf")
out_json = os.path.join(base, "chart_line_truth.json")

app = win32com.client.DispatchEx("PowerPoint.Application")
try:
    pres = app.Presentations.Open(pptx, ReadOnly=True, WithWindow=False)
    shp = pres.Slides(1).Shapes(1)
    ch = shp.Chart
    truth = {
        "shape": {
            "left": shp.Left, "top": shp.Top, "width": shp.Width, "height": shp.Height,
            "name": shp.Name, "type": shp.Type,
        },
        "chart": {
            "chart_type": ch.ChartType,
            "has_title": bool(ch.HasTitle),
            "has_legend": bool(ch.HasLegend),
        },
    }
    # legend position if any
    try:
        if ch.HasLegend:
            truth["chart"]["legend_position"] = int(ch.Legend.Position)
    except Exception as e:
        truth["chart"]["legend_err"] = str(e)
    pres.SaveAs(pdf, 32)
    pres.Close()
finally:
    app.Quit()

with open(out_json, "w", encoding="utf-8") as f:
    json.dump(truth, f, indent=2)
print(json.dumps(truth, indent=2))
print("pdf:", pdf, os.path.getsize(pdf))
