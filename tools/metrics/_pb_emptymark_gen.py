# -*- coding: utf-8 -*-
"""How tall is an EMPTY paragraph whose ¶-mark names an UNAVAILABLE font?

technical__002c6778's title block holds an empty paragraph whose mark rPr is
sz=13 (6.5pt) + rFonts ascii="Sabon LT Pro" (not installed).  Word gives that
line ~9.0pt, Oxi 7.55 — a 1.43pt seed that drifts down page 1 and flips a
keepNext pair at the page bottom by 0.13pt.

Each arm is a sandwich   [A<tag>] [mid] [Z<tag>]   on its own page, anchors in
Arial 10 with zero spacing, so

    H(mid) = (Z-A)_arm - (Z-A)_CTRL

is read straight off the PDF baselines (same size at both ends => the ascent
cancels).  Arms cover known fonts (model check), unavailable fonts (the real
case) and a size sweep, plus a VISIBLE run so the PDF span names the substitute.

  python _pb_emptymark_gen.py gen | bake | read
"""
import os, subprocess, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_emptymark")
DOCX = os.path.join(OUT, "emptymark.docx")
PDF = os.path.join(OUT, "emptymark.pdf")
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')

# (tag, kind, font, half-points)  kind: ctrl / empty / text
ARMS = [
    ("CTRL", "ctrl", None, None),
    ("ARI13", "empty", "Arial", 13),
    ("CAL13", "empty", "Calibri", 13),
    ("CAM13", "empty", "Cambria", 13),
    ("TNR13", "empty", "Times New Roman", 13),
    ("SGL13", "empty", "Segoe UI Light", 13),
    ("SAB13", "empty", "Sabon LT Pro", 13),
    ("MYR13", "empty", "Myriad Pro Light", 13),
    ("ZZZ13", "empty", "Nonexistent Face XQ", 13),
    ("SAB20", "empty", "Sabon LT Pro", 20),
    ("SAB40", "empty", "Sabon LT Pro", 40),
    ("SAB80", "empty", "Sabon LT Pro", 80),
    ("SABTX", "text", "Sabon LT Pro", 40),
    ("MYRTX", "text", "Myriad Pro Light", 40),
]


def anchor(tag, which):
    return ('<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>'
            '<w:t xml:space="preserve">%s%s</w:t></w:r></w:p>' % (which, tag))


def mid(kind, font, hp):
    if kind == "ctrl":
        return ""
    rpr = '<w:rFonts w:ascii="%s" w:hAnsi="%s"/><w:sz w:val="%d"/>' % (font, font, hp)
    if kind == "empty":
        return "<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr></w:p>" % rpr
    return ("<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr>"
            "<w:t>Hxg</w:t></w:r></w:p>" % (rpr, rpr))


SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for i, (tag, kind, font, hp) in enumerate(ARMS):
        body.append(anchor(tag, "A"))
        body.append(mid(kind, font, hp))
        body.append(anchor(tag, "Z"))
        if i != len(ARMS) - 1:
            body.append("<w:p><w:pPr>%s</w:pPr></w:p>" % SECT.replace("<w:sectPr>", '<w:sectPr><w:type w:val="nextPage"/>'))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           '><w:body>' + "".join(body) + SECT + "</w:body></w:document>")
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(ARMS), "arms")


def bake():
    import win32com.client as w
    app = w.DispatchEx("Word.Application"); app.Visible = False
    d = app.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(os.path.abspath(PDF), 17)
    finally:
        d.Close(False); app.Quit()
    print("baked", PDF)


def read():
    import fitz
    d = fitz.open(PDF)
    rows = {}
    for pi, pg in enumerate(d):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    t = "".join(c["c"] for c in s["chars"]).strip()
                    if t:
                        rows.setdefault(t, (s["origin"][1], s["font"], round(s["size"], 2), pi))
    ctrl = rows["ZCTRL"][0] - rows["ACTRL"][0]
    print("CTRL anchor-to-anchor = %.2f (Arial 10 line)\n" % ctrl)
    print("%-7s %-20s %5s %9s %9s   %s" % ("arm", "font", "pt", "gap", "H(mid)", "H/pt  | rendered font"))
    for tag, kind, font, hp in ARMS:
        if kind == "ctrl":
            continue
        a, z = rows.get("A" + tag), rows.get("Z" + tag)
        if not a or not z:
            print("%-7s MISSING" % tag); continue
        h = (z[0] - a[0]) - ctrl
        pt = hp / 2.0
        extra = ""
        if kind == "text":
            sp = rows.get("Hxg")
            extra = "  span=%s %.2f" % (sp[1], sp[2]) if sp else ""
        print("%-7s %-20s %5.1f %9.2f %9.2f   %6.4f%s" % (tag, font, pt, z[0] - a[0], h, h / pt, extra))


{"gen": gen, "bake": bake, "read": read}[sys.argv[1]]()
