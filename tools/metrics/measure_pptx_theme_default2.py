# measure_pptx_theme_default2.py - export + read span font per slide (theme modified: minor=TNR, major=Georgia)
import json, os, sys
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

OUT_DIR = os.path.abspath(r"pipeline_data\pptx_probes\theme_default2")
PPTX = os.path.join(OUT_DIR, "theme_default2.pptx")
PDF = os.path.join(OUT_DIR, "theme_default2.pdf")

app = win32com.client.DispatchEx("PowerPoint.Application")
try:
    pres = app.Presentations.Open(PPTX, WithWindow=False)
    pres.SaveAs(PDF, 32)
    pres.Close()
finally:
    app.Quit()
print("exported", PDF)

import fitz
doc = fitz.open(PDF)
result = []
for i, page in enumerate(doc):
    d = page.get_text("rawdict")
    row = {"slide": i+1, "spans": []}
    for blk in d["blocks"]:
        for line in blk["lines"]:
            for sp in line["spans"]:
                text = "".join(c["c"] for c in sp["chars"]).strip()
                if not text: continue
                row["spans"].append({"text": text, "font": sp["font"], "size": round(sp["size"],3)})
    result.append(row)
    print("S%02d" % (i+1), row["spans"])

out = os.path.join(OUT_DIR, "theme_measure.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(result, f, indent=1, ensure_ascii=False)
print("wrote", out)
