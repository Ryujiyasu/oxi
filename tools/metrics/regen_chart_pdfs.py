# -*- coding: utf-8 -*-
"""Regenerate chart1.pdf / chart2.pdf from the CURRENT pptx via PowerPoint COM,
then dump full text of each to confirm the 'Series 1' legend/title reproduces
(and chart2 shows nothing)."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

base = r"pipeline_data\pptx_probes"

app = win32com.client.DispatchEx("PowerPoint.Application")
for name in ("chart1", "chart2"):
    d = os.path.join(base, name)
    pptx = os.path.abspath(os.path.join(d, name + ".pptx"))
    pdf = os.path.abspath(os.path.join(d, name + "_regen.pdf"))
    pres = app.Presentations.Open(pptx, WithWindow=False)
    pres.SaveAs(pdf, 32)
    # read chart title / legend via COM
    try:
        sh = pres.Slides(1).Shapes(1)
        ch = sh.Chart
        info = {
            "has_title": ch.HasTitle,
            "has_legend": ch.HasLegend,
        }
        try:
            info["chart_title"] = ch.ChartTitle.TextFrame.TextRange.Text
        except Exception as e:
            info["chart_title_err"] = str(e)
        try:
            n = ch.Legend.LegendEntries().Count
            info["legend_entries"] = n
        except Exception as e:
            info["legend_err"] = str(e)
    except Exception as e:
        info = {"err": str(e)}
    pres.Close()
    print(f"=== {name} regen -> {pdf}")
    print(info)
app.Quit()
