# -*- coding: utf-8 -*-
"""Does a lead run holding only an INLINE IMAGE set the chars-indent unit?

`_pb_brrun_probe.py` settled the text-less lead runs Word counts (a page break
and a line break count, an empty run and a field-start run do not).  An inline
image is the remaining shape: it has no text but it is content.  The image is
copied verbatim out of a corpus document (media part + relationship), so Word
sees a real picture, and the stored twips are left deliberately stale.

Usage: python tools/metrics/_pb_imgrun_probe.py
"""
import glob
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
src = glob.glob("tools/golden-test/documents/docx/3a4f9fbe1a83*.docx")[0]
z = zipfile.ZipFile(src)
doc = z.read("word/document.xml").decode("utf-8")
head = doc[: doc.index("<w:body>") + len("<w:body>")]
sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)[-1]
sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
drawing = re.search(r"<w:drawing>\s*<wp:inline.*?</w:drawing>", doc, re.S).group(0)
# shrink the picture to one line's worth so it cannot dominate the line box
drawing = re.sub(r'<wp:extent cx="\d+" cy="\d+"/>', '<wp:extent cx="190500" cy="190500"/>', drawing)
drawing = re.sub(r'<a:ext cx="\d+" cy="\d+"/>', '<a:ext cx="190500" cy="190500"/>', drawing)
TEXT = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん" * 3


def rpr(sz):
    return ('<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (sz, sz))


def para(label, first_run, text_sz):
    ind = '<w:ind w:leftChars="100" w:left="210" w:hangingChars="100" w:hanging="200"/>'
    return "<w:p><w:pPr>%s</w:pPr>%s<w:r>%s<w:t>%s %s</w:t></w:r></w:p>" % (
        ind, first_run, rpr(text_sz), label, TEXT)


arms = [
    ("A", "<w:r>%s%s</w:r>" % (rpr(18), drawing), 28),   # inline image in a 9pt run, text 14pt
    ("B", "<w:r>%s%s</w:r>" % (rpr(28), drawing), 20),   # inline image in a 14pt run, text 10pt
    ("C", "", 28),                                        # control: text 14pt only
    ("D", "", 20),                                        # control: text 10pt only
]
paras = [para(label, first, sz) for label, first, sz in arms]
out = os.path.join("pipeline_data", "_pb_imgrun_probe.docx")
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
print("left margin %.2fpt | leftChars=100 hangingChars=100, stored hanging=200 (stale on purpose)" % margin)
for label, first, sz in arms:
    print("  %s lead=%s text=%.1fpt" % (label, "inline image in a %.1fpt run" % (int(re.search(r'w:sz w:val="(\d+)"', first).group(1)) / 2) if first else "none", sz / 2))
pd = fitz.open(pdf)
cur = None
for pno in range(len(pd)):
    for b in pd[pno].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if not t:
                continue
            if re.match(r"^[A-D] ", t):
                cur = t[0]
            print("p%d %s y=%6.1f x=%7.2f -> %5.0f tw  size %.1f  %s" % (
                pno + 1, cur, l["bbox"][1], l["bbox"][0], (l["bbox"][0] - margin) * 20, l["spans"][0]["size"], t[:12]))
