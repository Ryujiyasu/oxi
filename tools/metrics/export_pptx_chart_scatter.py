# -*- coding: utf-8 -*-
"""Export the SCATTER probe to PDF via PowerPoint COM and record the COM
truth (chart type / title / legend / series count) per slide."""
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
import win32com.client as win32

SRC = os.path.abspath(r"pipeline_data\pptx_probes\chart_scatter\chart_scatter.pptx")
PDF = os.path.abspath(r"pipeline_data\pptx_probes\chart_scatter\chart_scatter.pdf")
TRUTH = os.path.abspath(
    r"pipeline_data\pptx_probes\chart_scatter\chart_scatter_truth.json"
)


def main() -> None:
    app = win32.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(SRC, WithWindow=False)
        out = {"slide_w": pres.PageSetup.SlideWidth,
               "slide_h": pres.PageSetup.SlideHeight, "slides": []}
        for i in range(1, pres.Slides.Count + 1):
            sl = pres.Slides(i)
            rec = {"index": i, "shapes": []}
            for j in range(1, sl.Shapes.Count + 1):
                sh = sl.Shapes(j)
                d = {"name": sh.Name, "left": sh.Left, "top": sh.Top,
                     "width": sh.Width, "height": sh.Height}
                if sh.HasChart:
                    ch = sh.Chart
                    d["chart_type"] = ch.ChartType
                    d["has_title"] = bool(ch.HasTitle)
                    d["has_legend"] = bool(ch.HasLegend)
                    d["series_count"] = ch.SeriesCollection().Count
                rec["shapes"].append(d)
            out["slides"].append(rec)
            print(f"S{i}: " + " | ".join(
                f"{s.get('chart_type')} title={s.get('has_title')} "
                f"legend={s.get('has_legend')} ser={s.get('series_count')}"
                for s in rec["shapes"] if "chart_type" in s))
        pres.SaveAs(PDF, 32)
        pres.Close()
        json.dump(out, open(TRUTH, "w", encoding="utf-8"), indent=1)
        print("saved", PDF)
        print("saved", TRUTH)
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
