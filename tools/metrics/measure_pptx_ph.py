"""Measure placeholder position/size resolution truth from PowerPoint COM.

Reads ph_resolution.pptx, prints per-shape:
  slide, idx, Name, Type(MSO), Left, Top, Width, Height
so we can confirm (a) slide placeholders WITHOUT explicit xfrm take the
slideLayout's placeholder xfrm, and (b) slide placeholders WITH explicit
geometry keep their OWN geometry.
"""
import os
import sys
import json
sys.stdout.reconfigure(encoding="utf-8")

from win32com.client import DispatchEx

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "..", "pipeline_data", "pptx_probes", "ph_resolution.pptx")

app = DispatchEx("PowerPoint.Application")
pres = app.Presentations.Open(os.path.abspath(SRC), True, False, False)

rows = []
try:
    for si in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(si)
        for sh in slide.Shapes:
            rows.append({
                "slide": si,
                "idx": sh.Id,
                "name": sh.Name,
                "type": sh.Type,
                "left": round(sh.Left, 2),
                "top": round(sh.Top, 2),
                "width": round(sh.Width, 2),
                "height": round(sh.Height, 2),
            })
finally:
    pres.Close()
    app.Quit()

print(json.dumps(rows, ensure_ascii=False, indent=1))
