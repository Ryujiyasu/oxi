# -*- coding: utf-8 -*-
"""Does Word recompute w:left from w:leftChars, and with which unit?  The same
paragraph is repeated with the stored twips deliberately wrong, so the rendered
x says which of the two Word obeys.

Usage: python tools/metrics/_pb_leftchars_probe.py [host-docx-stem]
"""
import glob
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
stem = sys.argv[1] if len(sys.argv) > 1 else "*kaigo"
src = glob.glob("tools/golden-test/documents/docx/%s*.docx" % stem)[0]
z = zipfile.ZipFile(src)
doc = z.read("word/document.xml").decode("utf-8")
head = doc[: doc.index("<w:body>") + len("<w:body>")]
sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)[-1]
sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
TEXT = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん" * 2


def rpr(sz=None):
    s = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>'
    if sz:
        s += '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)
    return "<w:rPr>%s</w:rPr>" % s


arms = [
    ("A", '<w:ind w:leftChars="200" w:left="721"/>', 24),                      # kaigo's own pair (12pt run)
    ("B", '<w:ind w:leftChars="200" w:left="2000"/>', 24),                     # stored wildly wrong
    ("C", '<w:ind w:leftChars="200"/>', 24),                                   # chars only
    ("D", '<w:ind w:left="721"/>', 24),                                        # twips only
    ("E", '<w:ind w:leftChars="200" w:left="721"/>', 16),                      # same pair, 8pt run
    ("F", '<w:ind w:leftChars="200" w:left="2000"/>', 16),                     # 8pt run, stored wrong
]
paras = []
for label, ind, sz in arms:
    paras.append(
        "<w:p><w:pPr>%s</w:pPr><w:r>%s<w:t>%s %s</w:t></w:r></w:p>" % (ind, rpr(sz), label, TEXT)
    )
out = os.path.join("pipeline_data", "_pb_leftchars_probe.docx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as o:
    for item in z.infolist():
        if item.filename == "word/document.xml":
            o.writestr(item, (head + "".join(paras) + sect + "</w:body></w:document>").encode("utf-8"))
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

margin = int(re.search(r'w:left="(\d+)"', re.search(r"<w:pgMar [^>]*/>", sect).group(0)).group(1)) / 20.0
print("host %s | left margin %.2fpt" % (os.path.basename(src), margin))
for label, ind, sz in arms:
    print("  %s %s sz=%.1fpt" % (label, ind, sz / 2))
pd = fitz.open(pdf)
cur = None
for pno in range(len(pd)):
    for b in pd[pno].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if not t:
                continue
            if re.match(r"^[A-F] ", t):
                cur = t[0]
                print("p%d %s x=%7.2f -> %5.0f tw  %s" % (pno + 1, cur, l["bbox"][0], (l["bbox"][0] - margin) * 20, t[:10]))
