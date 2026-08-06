# -*- coding: utf-8 -*-
"""Convert chart_legend to PDF via PowerPoint COM and read chart COM props."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

def convert(pptx_path, pdf_path):
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = ppt.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        pres.SaveAs(os.path.abspath(pdf_path), 32)  # ppSaveAsPDF
        sh = pres.Slides(1).Shapes(1)
        info = {}
        if sh.HasChart:
            ch = sh.Chart
            info["has_title"] = bool(ch.HasTitle)
            info["has_legend"] = bool(ch.HasLegend)
            if ch.HasTitle:
                info["title"] = ch.ChartTitle.TextFrame.TextRange.Text
            try:
                leg = ch.Legend
                info["legend_position"] = int(leg.Position)
                info["legend_include"] = bool(leg.IncludeInLayout)
            except Exception as e:
                info["legend_err"] = str(e)
        pres.Close()
        print(pptx_path, "->", info)
    finally:
        ppt.Quit()

convert(r"pipeline_data\pptx_probes\chart_legend\chart_legend.pptx",
        r"pipeline_data\pptx_probes\chart_legend\chart_legend.pdf")
