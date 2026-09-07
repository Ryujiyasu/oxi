# -*- coding: utf-8 -*-
"""Which run's size does Word measure a chars indent on when the paragraph's
first run holds no text (a page break, an empty run, a bookmark)?  Hand-written
paragraphs are dropped into a corpus package (faithful styles/fonts), Word
exports the PDF, and every line's x is printed with the arm it belongs to.

Usage: python tools/metrics/_pb_brrun_probe.py [host-docx-stem]
"""
import glob
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
stem = sys.argv[1] if len(sys.argv) > 1 else "*8485f16"
src = glob.glob("tools/golden-test/documents/docx/%s*.docx" % stem)[0]
z = zipfile.ZipFile(src)
doc = z.read("word/document.xml").decode("utf-8")
head = doc[: doc.index("<w:body>") + len("<w:body>")]
sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)[-1]
sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)

TEXT = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん" * 3


def rpr(sz=None, sp=None):
    s = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>'
    if sp is not None:
        s += '<w:spacing w:val="%d"/>' % sp
    if sz is not None:
        s += '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)
    return "<w:rPr>%s</w:rPr>" % s


def para(label, first_run, text_sz, text_sp=None, hc=100, mark_sz=None):
    ind = '<w:ind w:leftChars="100" w:left="210" w:hangingChars="%d" w:hanging="%d"/>' % (hc, hc * 2)  # stale twips on purpose
    mark = "<w:rPr>%s</w:rPr>" % ('<w:sz w:val="%d"/>' % mark_sz if mark_sz else "")
    body = first_run + "<w:r>%s<w:t>%s %s</w:t></w:r>" % (rpr(text_sz, text_sp), label, TEXT)
    return "<w:p><w:pPr>%s%s</w:pPr>%s</w:p>" % (ind, mark, body)


arms = [
    # label, first run xml, text sz (half-pt)
    ("A", '<w:r>%s<w:br w:type="page"/></w:r>' % rpr(21), 24),            # break run 10.5pt, text 12pt
    ("B", '<w:r>%s<w:br w:type="page"/></w:r>' % rpr(18), 28),            # break run 9pt, text 14pt
    ("C", '<w:r>%s<w:br w:type="page"/></w:r>' % rpr(28), 20),            # break run 14pt, text 10pt
    ("D", '<w:r>%s</w:r>' % rpr(18), 28),                                  # empty run 9pt, text 14pt
    ("E", '<w:bookmarkStart w:id="7" w:name="bm"/><w:bookmarkEnd w:id="7"/>', 28),  # no run at all: text 14pt
    ("F", '<w:r>%s<w:br/></w:r>' % rpr(18), 28),                           # line-break run 9pt, text 14pt
    ("G", '<w:r>%s<w:br w:type="page"/></w:r>' % rpr(20, -20), 28),        # break run 10pt at -1pt tracking, text 14pt
    ("I", '<w:r>%s<w:tab/></w:r>' % rpr(18), 28),                          # tab-only run 9pt, text 14pt
    ("J", '<w:r>%s<w:fldChar w:fldCharType="begin"/></w:r>' % rpr(18), 28),  # field-start run 9pt, text 14pt
    ("K", '<w:r>%s<w:t xml:space="preserve"> </w:t></w:r>' % rpr(18), 28),  # space-only text run 9pt, text 14pt
]
paras = []
for label, first, tsz in arms:
    paras.append(para(label, first, tsz))
# H: text-first paragraph as control (14pt) so the arm math has a reference
paras.append(para("H", "", 28))
out = os.path.join("pipeline_data", "_pb_brrun_probe.docx")
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

pd = fitz.open(pdf)
margin = int(re.search(r'w:left="(\d+)"', re.search(r"<w:pgMar [^>]*/>", sect).group(0)).group(1)) / 20.0
print("left margin %.2fpt | stored hanging twips are hc*2 (deliberately stale); leftChars=100" % margin)
cur = None
for pno in range(len(pd)):
    for b in pd[pno].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if not t:
                continue
            if re.match(r"^[A-K] ", t):
                cur = t[0]
            print("p%d %s y=%6.1f x=%7.2f (+%6.2fpt = %4.0f tw) size %.1f %s" % (pno + 1, cur, l["bbox"][1], l["bbox"][0], l["bbox"][0] - margin, (l["bbox"][0] - margin) * 20, l["spans"][0]["size"], t[:12]))
