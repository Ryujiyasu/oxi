# -*- coding: utf-8 -*-
"""Chart spec wave-1 oracle: open chart1.pptx in PowerPoint COM, record chart
shape geometry + axis/data-label info, export PDF for fitz measurement."""
import sys, os, json
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

base = r"pipeline_data\pptx_probes\chart1"
pptx_path = os.path.abspath(os.path.join(base, "chart1.pptx"))
pdf_path = os.path.abspath(os.path.join(base, "chart1.pdf"))

app = win32com.client.DispatchEx("PowerPoint.Application")
pres = app.Presentations.Open(pptx_path, WithWindow=False)

sh = pres.Slides(1).Shapes(1)
data = {
    "shape": {
        "left": sh.Left, "top": sh.Top, "width": sh.Width, "height": sh.Height,
        "name": sh.Name, "type": sh.Type,
        "has_chart": sh.HasChart,
    },
}
if sh.HasChart:
    ch = sh.Chart
    data["chart"] = {
        "type": ch.ChartType,
        "has_title": ch.HasTitle,
        "chart_area": {"left": ch.ChartArea.Left, "top": ch.ChartArea.Top,
                       "width": ch.ChartArea.Width, "height": ch.ChartArea.Height},
        "has_legend": ch.HasLegend,
    }
    try:
        data["chart"]["legend"] = {"left": ch.Legend.Left, "top": ch.Legend.Top,
                                   "width": ch.Legend.Width, "height": ch.Legend.Height}
    except Exception as e:
        data["chart"]["legend_err"] = str(e)

with open(os.path.join(base, "chart1_truth.json"), "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=1)

pres.SaveAs(pdf_path, 32)
pres.Close()
app.Quit()
print("truth saved")
