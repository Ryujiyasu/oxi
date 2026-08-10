# -*- coding: utf-8 -*-
"""Measure chart_title_line / chart_title_line2 / chart_title_pie /
chart_title_pie2: open in PowerPoint COM, read the chart title props
(text / geometry / font), export the PDF for fitz measurement."""
import sys, os, json
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client


def measure(base_name):
    base = os.path.join("pipeline_data", "pptx_probes", base_name)
    pptx_path = os.path.abspath(os.path.join(base, base_name + ".pptx"))
    pdf_path = os.path.abspath(os.path.join(base, base_name + ".pdf"))

    app = win32com.client.DispatchEx("PowerPoint.Application")
    pres = app.Presentations.Open(pptx_path, WithWindow=False)
    data = {}
    for si in range(1, pres.Slides.Count + 1):
        sh = pres.Slides(si).Shapes(1)
        d = {
            "shape": {
                "left": sh.Left, "top": sh.Top, "width": sh.Width, "height": sh.Height,
                "name": sh.Name, "type": sh.Type, "has_chart": sh.HasChart,
            },
        }
        if sh.HasChart:
            ch = sh.Chart
            d["chart"] = {
                "type": ch.ChartType,
                "has_title": bool(ch.HasTitle),
                "has_legend": bool(ch.HasLegend),
            }
            if ch.HasTitle:
                t = ch.ChartTitle
                title_info = {
                    "left": t.Left, "top": t.Top,
                    "width": t.Width, "height": t.Height,
                }
                try:
                    title_info["text"] = t.Caption
                except Exception:
                    try:
                        title_info["text"] = t.TextFrame2.TextRange.Text
                    except Exception as e:
                        title_info["text_err"] = str(e)
                try:
                    f = t.Format.TextFrame2.TextRange.Font
                    title_info["format"] = {
                        "size": f.Size, "name": f.Name, "bold": int(f.Bold),
                    }
                except Exception as e:
                    title_info["format_err"] = str(e)
                d["chart"]["title"] = title_info
        data["slide%d" % si] = d
    with open(os.path.join(base, base_name + "_truth.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    pres.SaveAs(pdf_path, 32)
    pres.Close()
    app.Quit()
    print(base_name, "truth saved")


for name in ("chart_title_line", "chart_title_line2", "chart_title_pie", "chart_title_pie2"):
    measure(name)
