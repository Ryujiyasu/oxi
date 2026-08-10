# -*- coding: utf-8 -*-
"""Chart data-labels oracle: open chart_datalabel.pptx in PowerPoint COM,
record each slide's chart/data-label properties, export PDF for fitz
measurement (associating label text spans with bar geometry)."""
import sys, os, json
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

base = r"pipeline_data\pptx_probes\chart_datalabel"
pptx_path = os.path.abspath(os.path.join(base, "chart_datalabel.pptx"))
pdf_path = os.path.abspath(os.path.join(base, "chart_datalabel.pdf"))

app = win32com.client.DispatchEx("PowerPoint.Application")
pres = app.Presentations.Open(pptx_path, WithWindow=False)

out = {"slides": []}
for i in range(1, pres.Slides.Count + 1):
    sh = pres.Slides(i).Shapes(1)
    s = {"shape": {
        "left": sh.Left, "top": sh.Top, "width": sh.Width, "height": sh.Height,
        "name": sh.Name, "type": sh.Type, "has_chart": sh.HasChart}}
    if sh.HasChart:
        ch = sh.Chart
        chart = {"type": ch.ChartType, "has_title": ch.HasTitle,
                 "chart_area": {"left": ch.ChartArea.Left, "top": ch.ChartArea.Top,
                                "width": ch.ChartArea.Width, "height": ch.ChartArea.Height}}
        try:
            plot = ch.PlotArea
            chart["plot_area"] = {"left": plot.Left, "top": plot.Top,
                                  "width": plot.Width, "height": plot.Height}
        except Exception as e:
            chart["plot_area_err"] = str(e)
        # Series-level data labels
        try:
            labels = []
            for si in range(1, ch.SeriesCollection().Count + 1):
                ser = ch.SeriesCollection().Item(si)
                entry = {}
                try:
                    dl = ser.DataLabels()
                    entry["has_data_labels"] = dl.Count > 0
                    if dl.Count > 0:
                        one = dl.Item(1)
                        try:
                            entry["position"] = one.Position
                        except Exception as e:
                            entry["position_err"] = str(e)
                        try:
                            entry["number_format"] = one.NumberFormat
                        except Exception as e:
                            entry["numfmt_err"] = str(e)
                        try:
                            entry["show_value"] = one.ShowValue
                        except Exception as e:
                            entry["showval_err"] = str(e)
                except Exception as e:
                    entry["dl_err"] = str(e)
                labels.append(entry)
            chart["series_labels"] = labels
        except Exception as e:
            chart["series_labels_err"] = str(e)
        s["chart"] = chart
    out["slides"].append(s)

with open(os.path.join(base, "chart_datalabel_truth.json"), "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=1)

pres.SaveAs(pdf_path, 32)
pres.Close()
app.Quit()
print("truth saved + pdf exported")
