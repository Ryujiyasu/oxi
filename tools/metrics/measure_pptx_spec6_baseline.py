# measure_pptx_spec6_baseline.py — Word: export spec6_baseline.pptx to PDF, read first-line baseline per slide.
# Then compute the first-baseline offset from the text-area top (=72+3.6=75.6) as an em multiple.
import json
import os
import sys
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

OUT_DIR = r"C:\Users\ryuji\oxi-main\pipeline_data\pptx_probes\spec6_baseline"
PPTX = os.path.join(OUT_DIR, "spec6_baseline.pptx")
PDF = os.path.join(OUT_DIR, "spec6_baseline.pdf")

app = win32com.client.DispatchEx("PowerPoint.Application")
try:
    pres = app.Presentations.Open(PPTX, WithWindow=False)
    pres.SaveAs(PDF, 32)  # ppSaveAsPDF
    pres.Close()
finally:
    app.Quit()
print("exported", PDF)

import fitz
doc = fitz.open(PDF)
text_area_top = 72.0 + 3.6  # shape top 72 + margin_top 0.05in=3.6pt
result = []
for i, page in enumerate(doc):
    d = page.get_text("rawdict")
    # first span on the page
    blk = d["blocks"][0]
    line = blk["lines"][0]
    sp = line["spans"][0]
    baseline = sp["origin"][1]
    x0 = sp["origin"][0]
    text = "".join(c["c"] for c in sp["chars"]).strip()
    # font size from the span (name carries the label)
    fs = float(sp["size"])
    result.append({
        "slide": i + 1,
        "text": text,
        "fs": fs,
        "baseline": round(baseline, 3),
        "offset_pt": round(baseline - text_area_top, 3),
        "offset_em": round((baseline - text_area_top) / fs, 6),
        "ascender": round(sp["ascender"], 6),
    })

out = os.path.join(OUT_DIR, "baseline_measure.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump({"text_area_top": text_area_top, "results": result}, f, indent=1, ensure_ascii=False)
print("wrote", out)
for r in result:
    print(r)
