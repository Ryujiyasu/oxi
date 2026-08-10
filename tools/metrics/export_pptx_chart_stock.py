# -*- coding: utf-8 -*-
"""Export the stock probe deck to PDF via PowerPoint COM."""
import os
import sys

import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.abspath(r"pipeline_data\pptx_probes\chart_stock")
PPTX = os.path.join(ROOT, "chart_stock.pptx")
PDF = os.path.join(ROOT, "chart_stock.pdf")


def main():
    app = win32.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(PPTX, WithWindow=False)
        try:
            for i, sl in enumerate(pres.Slides, 1):
                for sh in sl.Shapes:
                    if sh.HasChart:
                        c = sh.Chart
                        print("slide %d chart_type=%s has_title=%s has_legend=%s"
                              % (i, c.ChartType, c.HasTitle, c.HasLegend))
            pres.SaveAs(PDF, 32)
        finally:
            pres.Close()
    finally:
        app.Quit()
    print("wrote", PDF)


if __name__ == "__main__":
    main()
