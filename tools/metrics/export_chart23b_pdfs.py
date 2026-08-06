# -*- coding: utf-8 -*-
"""Convert chart2b/chart3 to PDF via PowerPoint COM and read title/legend COM values."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

def convert(pptx_path, pdf_path):
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = ppt.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        pres.SaveAs(os.path.abspath(pdf_path), 32)  # ppSaveAsPDF
        # read chart COM props
        sh = pres.Slides(1).Shapes(1)
        info = {}
        if sh.HasChart:
            ch = sh.Chart
            try:
                info["has_title"] = bool(ch.HasTitle)
            except Exception as e:
                info["has_title_err"] = str(e)
            try:
                info["has_legend"] = bool(ch.HasLegend)
            except Exception as e:
                info["has_legend_err"] = str(e)
            if ch.HasTitle:
                try:
                    info["title"] = ch.ChartTitle.TextFrame.TextRange.Text
                except Exception as e:
                    info["title_err"] = str(e)
        pres.Close()
        print(pptx_path, "->", info)
    finally:
        ppt.Quit()

convert(r"pipeline_data\pptx_probes\chart2b\chart2b.pptx",
        r"pipeline_data\pptx_probes\chart2b\chart2b.pdf")
convert(r"pipeline_data\pptx_probes\chart3\chart3.pptx",
        r"pipeline_data\pptx_probes\chart3\chart3.pdf")
