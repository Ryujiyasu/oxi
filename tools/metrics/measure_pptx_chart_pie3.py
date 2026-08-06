#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""COM-export chart_pie3.pptx to PDF + record per-slide truth."""
import json
import os

import win32com.client

import os

BASE = r"C:\Users\ryuji\oxi-main\pipeline_data\pptx_probes\chart_pie3"
PPTX = os.path.join(BASE, "chart_pie3.pptx")
PDF = os.path.join(BASE, "chart_pie3.pdf")
OUT_JSON = os.path.join(BASE, "chart_pie3_truth.json")


def main():
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(PPTX, WithWindow=False)
        pres.SaveAs(PDF, 32)  # ppSaveAsPDF
        truth = []
        for sl in pres.Slides:
            for sh in sl.Shapes:
                if not sh.HasChart:
                    continue
                ch = sh.Chart
                hl = bool(ch.HasLegend)
                rec = {
                    "chart_type": ch.ChartType,
                    "has_title": bool(ch.HasTitle),
                    "has_legend": hl,
                    "series_count": ch.SeriesCollection().Count,
                }
                if hl:
                    leg = ch.Legend
                    rec["legend_include"] = bool(leg.IncludeInLayout)
                    rec["legend_pos"] = leg.Position
                truth.append(rec)
        pres.Close()
    finally:
        app.Quit()
    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(truth, f, ensure_ascii=False, indent=1)
    print(json.dumps(truth, ensure_ascii=False))


if __name__ == "__main__":
    main()
