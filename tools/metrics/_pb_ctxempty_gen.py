# -*- coding: utf-8 -*-
"""Does contextualSpacing suppress the after-spacing between EMPTY paragraphs?

correspondence__000f9471's page 1 holds four ListParagraph (contextualSpacing)
EMPTY spacer paragraphs.  Oxi suppresses the 10pt docDefaults after between
them (S657); Word's page-2 geometry only closes if at least one of those gaps
KEEPS its 10pt.  This probe measures the gap directly for

    T : ctx style, both paragraphs carry TEXT      (the S657 derivation shape)
    E : ctx style, the paragraph BETWEEN is EMPTY
    N : ctx style + numPr on both                  (real list items)
    C : plain Normal control (no contextualSpacing)

Each arm is a sandwich  [A<tag>] [mid] [Z<tag>]  so the gap is read from the
PDF baselines of the two markers: gap(A->Z) - line = the spacing that survived.

  python _pb_ctxempty_gen.py gen | bake | read
"""
import os, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_ctxempty")
DOCX = os.path.join(OUT, "ctxempty.docx")
PDF = os.path.join(OUT, "ctxempty.pdf")

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>'
          '<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style></w:styles>')
NUM = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + NS + '>'
       '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
       '<w:numFmt w:val="bullet"/><w:lvlText w:val="-"/><w:lvlJc w:val="left"/>'
       '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
       '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')

def para(style, text, num=False):
    pr = "<w:pPr>"
    if style:
        pr += '<w:pStyle w:val="%s"/>' % style
    if num:
        pr += '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
    pr += "</w:pPr>"
    r = ('<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % text) if text else ""
    return "<w:p>" + pr + r + "</w:p>"

ARMS = [("T", "ListParagraph", "mid", False),
        ("E", "ListParagraph", "",    False),
        ("N", "ListParagraph", "mid", True),
        ("C", None,            "mid", False),
        ("D", None,            "",    False)]

def build():
    b = []
    for tag, st, mid, num in ARMS:
        b.append(para(st, "AA" + tag, num))
        b.append(para(st, mid, num))
        b.append(para(st, "ZZ" + tag, num))
        b.append(para(None, "gap" + tag))
    body = "".join(b) + ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                         '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
            '><w:body>' + body + '</w:body></w:document>')

def gen():
    os.makedirs(OUT, exist_ok=True)
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/document.xml", build())
        z.writestr("word/styles.xml", STYLES); z.writestr("word/numbering.xml", NUM)
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
    doc = fitz.open(PDF); rows = []
    for pi, pg in enumerate(doc):
        r = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    t = "".join(c["c"] for c in s["chars"])
                    if not t.strip():
                        continue
                    y = round(s["origin"][1], 2)
                    k = next((k for k in r if abs(k - y) <= 0.75), y)
                    r.setdefault(k, []).append((s["origin"][0], t))
        for y in sorted(r):
            rows.append((pi, y, "".join(t for _, t in sorted(r[y])).strip()))
    idx = {t: (p, y) for p, y, t in rows}
    print("%-4s %8s %8s %8s   %s" % ("arm", "A->Z", "line", "spacing", "note"))
    for tag, st, mid, num in ARMS:
        a, z = idx.get("AA" + tag), idx.get("ZZ" + tag)
        if not a or not z:
            print("%-4s missing" % tag); continue
        span = z[1] - a[1]
        n_lines = 2 if mid else 1
        line = 13.8
        print("%-4s %8.2f %8.2f %8.2f   mid=%r num=%s style=%s"
              % (tag, span, line, span - n_lines * line, mid, num, st))

if __name__ == "__main__":
    {"gen": gen, "bake": bake, "read": read}[sys.argv[1]]()
