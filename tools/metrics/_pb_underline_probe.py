# -*- coding: utf-8 -*-
"""Where Word draws a single underline, as a function of font and size.

One paragraph per arm, each a short underlined run in a named font at a named
size, dropped into a corpus package so Word sees real styles.  The PDF is read
back with fitz: for every arm the rule's offset below the run's baseline and its
thickness are printed in points and as a fraction of the size, which is what a
renderer needs to place it.

Usage: python tools/metrics/_pb_underline_probe.py [host-docx-stem]
"""
import glob
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
stem = sys.argv[1] if len(sys.argv) > 1 else "1ec1"
src = glob.glob("tools/golden-test/documents/docx/%s*.docx" % stem)[0]
z = zipfile.ZipFile(src)
doc = z.read("word/document.xml").decode("utf-8")
head = doc[: doc.index("<w:body>") + len("<w:body>")]
sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)[-1]
sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)

FONTS = os.environ.get("UL_FONTS", "ＭＳ 明朝,ＭＳ ゴシック,游明朝,游ゴシック,メイリオ,Century").split(",")
SIZES = [float(x) for x in os.environ.get("UL_SIZES", "8,9,10,10.5,12,14,20").split(",")]
# UL_TEXT=latin puts LATIN text in every arm (same font, other script) -- the
# discriminator test for "does the run's script move Word's underline?"
_MODE = os.environ.get("UL_TEXT", "auto")
TEXT = {"Century": "Underline sample", None: "下線の見本あいう"}
if _MODE == "latin":
    TEXT = {None: "Underline sample"}
elif _MODE == "cjk":
    TEXT = {None: "下線の見本あいう"}

paras = []
arms = []
for fi, font in enumerate(FONTS):
    for si, size in enumerate(SIZES):
        label = "%s%s" % (chr(ord("A") + fi), si)
        body = TEXT.get(font, TEXT[None])
        # UL_KIND=double measures the two rules of a double underline
        kind = os.environ.get("UL_KIND", "single")
        rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
               '<w:sz w:val="%d"/><w:szCs w:val="%d"/><w:u w:val="%s"/></w:rPr>'
               % (font, font, font, round(size * 2), round(size * 2), kind))
        lead = ('<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>'
                '<w:sz w:val="20"/></w:rPr>')
        paras.append('<w:p><w:pPr><w:spacing w:line="480" w:lineRule="exact"/></w:pPr>'
                     '<w:r>%s<w:t xml:space="preserve">%s </w:t></w:r>'
                     '<w:r>%s<w:t>%s</w:t></w:r></w:p>' % (lead, label, rpr, body))
        arms.append((label, font, size))
out = os.path.join("pipeline_data", "_pb_underline_probe.docx")
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
found = {}
for pno in range(len(pd)):
    page = pd[pno]
    spans = [s for b in page.get_text("dict")["blocks"] for l in b.get("lines", []) for s in l["spans"] if s["text"].strip()]
    rules = [dr["rect"] for dr in page.get_drawings() if dr["rect"].height <= 2.5 and dr["rect"].width >= 5]
    for r in rules:
        best = None
        for s in spans:
            base = s["origin"][1]
            if base > r.y0 + 0.6 or base < r.y0 - 6.0:
                continue
            ox = min(s["bbox"][2], r.x1) - max(s["bbox"][0], r.x0)
            if ox < 0.6 * (r.x1 - r.x0):
                continue
            d = r.y0 - base
            if best is None or d < best[0]:
                best = (d, s)
        if best is None:
            continue
        d, s = best
        # the arm label is the first span on the same line
        line_spans = [x for x in spans if abs(x["origin"][1] - s["origin"][1]) < 0.6]
        lab = line_spans[0]["text"].strip().split()[0] if line_spans else "?"
        prev = found.get(lab)
        if prev is None or d < prev[2]:
            found[lab] = (s["font"], round(s["size"], 2), round(d, 3), round(r.height, 3),
                          prev[2] if prev else None)
        else:
            found[lab] = prev[:4] + (round(d, 3),)
print("arm  requested font          size   pdf font              dy(pt)  thick  dy/size  th/size")
for lab, font, size in arms:
    if lab not in found:
        print("  %-4s %-22s %5.1f   (no rule found)" % (lab, font, size))
        continue
    pf, ps, dy, th, dy2 = found[lab]
    print("  %-4s %-22s %5.1f   %-20s %6.3f %6.3f  %7.4f %7.4f  %s"
          % (lab, font, size, pf, dy, th, dy / ps, th / ps,
             ("2nd %6.3f (gap %.3f, /size %.4f)" % (dy2, dy2 - dy, dy2 / ps)) if dy2 else ""))
