# -*- coding: utf-8 -*-
"""Faithful slice of a corpus docx keeping only the paragraphs whose XML matches a
regex, exported through Word to PDF, then every line's x/y printed next to the
paragraph's <w:ind>: the instrument that asks Word which indent it renders when
the stored twips and the chars attributes disagree.

Usage: python tools/metrics/_pb_ind_slice.py <docx-stem> <regex> [max-paras]
"""
import glob
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
stem, pat = sys.argv[1], sys.argv[2]
limit = int(sys.argv[3]) if len(sys.argv) > 3 else 6
src = glob.glob("tools/golden-test/documents/docx/%s*.docx" % stem)[0]
z = zipfile.ZipFile(src)
doc = z.read("word/document.xml").decode("utf-8")
head = doc[: doc.index("<w:body>") + len("<w:body>")]
sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)[-1]
sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
paras = re.findall(r"<w:p[ >].*?</w:p>", doc, re.S)
keep = [p for p in paras if re.search(pat, p)][:limit]
if os.environ.get("SLICE_STRETCH"):
    # repeat the first run's text so a one-line paragraph wraps and shows its continuation indent
    n = int(os.environ["SLICE_STRETCH"])
    keep = [re.sub(r"(<w:t[^>]*>)([^<]+)(</w:t>)", lambda m: m.group(1) + m.group(2) * n + m.group(3), p) for p in keep]
print("kept", len(keep), "of", len(paras), "| sect", re.sub(r' w:rsid\w*="[^"]*"', "", sect)[:300])
out = os.path.join("pipeline_data", "_pb_ind_slice_%s.docx" % re.sub(r"[^A-Za-z0-9_]", "", stem))
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as o:
    for item in z.infolist():
        if item.filename == "word/document.xml":
            o.writestr(item, (head + "".join(keep) + sect + "</w:body></w:document>").encode("utf-8"))
        elif item.filename.startswith("word/header") or item.filename.startswith("word/footer"):
            continue
        else:
            data = z.read(item.filename)
            if item.filename == "word/_rels/document.xml.rels":
                data = re.sub(rb'<Relationship [^>]*Target="(header|footer)\d*\.xml"[^>]*/>', b"", data)
            if item.filename == "[Content_Types].xml":
                data = re.sub(rb'<Override [^>]*PartName="/word/(header|footer)\d*\.xml"[^>]*/>', b"", data)
            o.writestr(item, data)
import win32com.client as w  # noqa: E402

app = w.Dispatch("Word.Application")
app.Visible = False
pdf = os.path.abspath(out)[:-5] + ".pdf"
try:
    d = app.Documents.Open(os.path.abspath(out), ReadOnly=True)
    d.ExportAsFixedFormat(pdf, 17)
    d.Close(False)
finally:
    app.Quit()
import fitz  # noqa: E402

for i, p in enumerate(keep):
    ind = re.search(r"<w:ind [^>]*/>", p)
    r1 = re.search(r"<w:r>.*?</w:r>|<w:r [^>]*>.*?</w:r>", re.sub(r"<w:pPr>.*?</w:pPr>", "", p, flags=re.S), re.S)
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", r1.group(0), re.S).group(0) if r1 and "<w:rPr>" in r1.group(0) else ""
    txt = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p))[:16]
    print("para%d %s | run1 %s | %s" % (i + 1, re.sub(r' w:rsid\w*="[^"]*"', "", ind.group(0)) if ind else "-", re.sub(r"\s+", " ", rpr)[:160], txt))
pd = fitz.open(pdf)
for pno in range(len(pd)):
    for b in pd[pno].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if t:
                print("Word p%d y=%6.1f x=%7.2f x1=%6.1f %s" % (pno + 1, l["bbox"][1], l["bbox"][0], l["bbox"][2], t[:30]))
