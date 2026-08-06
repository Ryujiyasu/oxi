#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Export chart_pie.pptx to PDF via PowerPoint COM + dump chart COM properties.

Outputs:
  pipeline_data/pptx_probes/chart_pie/chart_pie.pdf
  pipeline_data/pptx_probes/chart_pie/chart_pie_truth.json
"""
import json
import os
import sys

import win32com.client

SRC = os.path.abspath(r"pipeline_data\pptx_probes\chart_pie\chart_pie.pptx")
PDF = os.path.abspath(r"pipeline_data\pptx_probes\chart_pie\chart_pie.pdf")
OUT = os.path.abspath(r"pipeline_data\pptx_probes\chart_pie\chart_pie_truth.json")


def main():
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(SRC, WithWindow=False)
        s1 = pres.Slides(1)
        shp = s1.Shapes(1)
        assert shp.HasChart, "shape 1 is not a chart"
        ch = shp.Chart
        truth = {
            "shape_left": shp.Left / 12700.0,
            "shape_top": shp.Top / 12700.0,
            "shape_w": shp.Width / 12700.0,
            "shape_h": shp.Height / 12700.0,
            "chart_type": ch.ChartType,
            "has_title": ch.HasTitle,
            "has_legend": ch.HasLegend,
            "legend_include": ch.Legend.IncludeInLayout if ch.HasLegend else None,
            "legend_pos": ch.Legend.Position if ch.HasLegend else None,
            "series_count": ch.SeriesCollection().Count,
        }
        pres.SaveAs(PDF, 32)  # ppSaveAsPDF
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(truth, f, ensure_ascii=False, indent=2)
        print(json.dumps(truth, ensure_ascii=False))
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
