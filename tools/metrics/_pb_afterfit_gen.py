# -*- coding: utf-8 -*-
"""Must a paragraph's space-AFTER fit on the page for the paragraph to stay?

Two docs now push a paragraph that fits the content bottom by <1pt while its
after-spacing would hang past it (policies__0028d1be's CORE COMPETENCIES,
correspondence__000f9471's last empty).  If Word required
    box_bottom + space_after <= content_bottom
the page capacity would drop by exactly `after` - measurable as the last
baseline on page 1 for a uniform filler.

  python _pb_afterfit_gen.py gen | bake | read
"""
import os, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_afterfit")
DOCX = os.path.join(OUT, "afterfit.docx")
PDF = os.path.join(OUT, "afterfit.pdf")
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
RPR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>'
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>' + RPR + '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>')
ARMS = [("A", 0), ("B", 200), ("C", 400)]
NLINE = 70

def sect(last):
    sp = ('<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="720" '
          'w:bottom="720" w:left="720" w:header="708" w:footer="708"/>')
    return sp if last else "<w:p><w:pPr><w:sectPr>" + sp + "</w:sectPr></w:pPr></w:p>"

def build():
    b = []
    for i, (tag, after) in enumerate(ARMS):
        for n in range(NLINE):
            b.append('<w:p><w:pPr><w:spacing w:before="0" w:after="%d" w:line="240" '
                     'w:lineRule="auto"/></w:pPr><w:r><w:t xml:space="preserve">%s%02d wwww</w:t></w:r></w:p>'
                     % (after, tag, n))
        b.append(sect(i == len(ARMS) - 1) if i < len(ARMS) - 1 else "<w:sectPr>" + sect(True) + "</w:sectPr>")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
            '><w:body>' + "".join(b) + '</w:body></w:document>')

def gen():
    os.makedirs(OUT, exist_ok=True)
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/document.xml", build()); z.writestr("word/styles.xml", STYLES)
    print("generated", os.path.abspath(DOCX))

def bake():
    import win32com.client as wc
    app = wc.DispatchEx("Word.Application"); app.Visible = False; app.DisplayAlerts = 0
    try:
        d = app.Documents.Open(os.path.abspath(DOCX), ReadOnly=True, AddToRecentFiles=False)
        d.ExportAsFixedFormat(os.path.abspath(PDF), 17); d.Close(0)
    finally:
        app.Quit()
    print("baked", os.path.abspath(PDF))

def read():
    import fitz
    doc = fitz.open(PDF); per = {}
    for pi, pg in enumerate(doc):
        rows = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    t = "".join(c["c"] for c in s["chars"])
                    if not t.strip():
                        continue
                    y = round(s["origin"][1], 2)
                    k = next((k for k in rows if abs(k - y) <= 0.75), y)
                    rows.setdefault(k, []).append(t)
        for y in sorted(rows):
            per.setdefault("".join(rows[y]).strip()[:3], (pi, y))
    print("%-4s %6s | %8s %8s %5s | %s" % ("arm", "after", "firstY", "lastY", "n", "box_bottom(last)+after"))
    for tag, after in ARMS:
        ys = [(p, y) for n in range(NLINE) for (p, y) in [per.get("%s%02d" % (tag, n), (None, None))] if p is not None]
        if not ys:
            print("%-4s missing" % tag); continue
        p0 = ys[0][0]; f = [y for p, y in ys if p == p0]
        bb = f[-1] + 2.596
        print("%-4s %6.1f | %8.2f %8.2f %5d | %8.2f  %8.2f" % (tag, after / 20.0, f[0], f[-1], len(f), bb, bb + after / 20.0))

if __name__ == "__main__":
    {"gen": gen, "bake": bake, "read": read}[sys.argv[1]]()
