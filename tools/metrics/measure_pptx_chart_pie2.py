#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Export chart_pie2.pptx (4 slides) to PDF via PowerPoint COM + dump chart props."""
import json
import os

import win32com.client

SRC = os.path.abspath(r"pipeline_data\pptx_probes\chart_pie2\chart_pie2.pptx")
PDF = os.path.abspath(r"pipeline_data\pptx_probes\chart_pie2\chart_pie2.pdf")
OUT = os.path.abspath(r"pipeline_data\pptx_probes\chart_pie2\chart_pie2_truth.json")


def main():
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(SRC, WithWindow=False)
        slides = []
        for si in range(1, pres.Slides.Count + 1):
            shp = pres.Slides(si).Shapes(1)
            ch = shp.Chart
            slides.append({
                "slide": si,
                "chart_type": ch.ChartType,
                "has_title": bool(ch.HasTitle),
                "has_legend": bool(ch.HasLegend),
                "legend_include": bool(ch.Legend.IncludeInLayout) if ch.HasLegend else None,
                "legend_pos": int(ch.Legend.Position) if ch.HasLegend else None,
                "series_count": ch.SeriesCollection().Count,
            })
        pres.SaveAs(PDF, 32)
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(slides, f, ensure_ascii=False, indent=2)
        print(json.dumps(slides, ensure_ascii=False))
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
