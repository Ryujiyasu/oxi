# measure_pptx_theme_default.py - Word: export theme_default.pptx to PDF, read each slide's span font.
# Also read theme1.xml minorFont/majorFont latin typeface from the pptx zip.
import json, os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

OUT_DIR = r"C:\Users\ryuji\oxi-main\pipeline_data\pptx_probes\theme_default"
PPTX = os.path.join(OUT_DIR, "theme_default.pptx")
PDF = os.path.join(OUT_DIR, "theme_default.pdf")

# theme fonts from the pptx zip
def theme_fonts(path):
    out = {}
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if n.startswith("ppt/theme/") and n.endswith(".xml")]
        for n in names:
            xml = z.read(n).decode("utf-8", "replace")
            def grab(tag):
                i = xml.find("<a:%s>" % tag)
                if i < 0: return None
                j = xml.find("<a:latin", i)
                if j < 0: return None
                k = xml.find("typeface=", j)
                if k < 0: return None
                s = xml.find('"', k); e = xml.find('"', s+1)
                return xml[s+1:e]
            out[n] = {"minor": grab("minorFont"), "major": grab("majorFont")}
    return out

tf = theme_fonts(PPTX)
print("THEME", json.dumps(tf, ensure_ascii=False))

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
    for blk in d["blocks"]:
        for line in blk["lines"]:
            for sp in line["spans"]:
                text = "".join(c["c"] for c in sp["chars"]).strip()
                if not text: continue
                result.append({"slide": i+1, "text": text, "font": sp["font"],
                               "size": round(sp["size"],3)})
    # keep first non-empty per slide only
    seen = [r for r in result if r["slide"]==i+1]
    print("S%02d" % (i+1), seen[-1] if seen else "EMPTY")
out = os.path.join(OUT_DIR, "theme_measure.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump({"theme": tf, "spans": result}, f, indent=1, ensure_ascii=False)
print("wrote", out)
