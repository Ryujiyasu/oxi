# -*- coding: utf-8 -*-
"""What does Word require of a page-bottom EMPTY paragraph?

S736 grants an empty paragraph a 2.5pt tolerance past the content bottom, but
no constant serves every specimen: Word KEEPS probexempty at +2.0 and
reference__0042471c at +2.5 overflow, yet PUSHES correspondence__000f9471's
last spacer whose box FITS by 0.72pt.  This probe settles the rule directly.

Arms sweep the BOTTOM MARGIN so the trailing empty paragraph's box crosses the
content bottom, and read that paragraph's page with COM.  The flip point
separates

    A: the empty's BOX must fit                 (flip at cbot = box bottom)
    B: the empty's box AND its after must fit   (flip 10pt earlier)

Both candidates are inside the swept range.  A TEXT control arm re-checks the
_pb_afterfit result (a text paragraph's after may hang past the bottom).

  python _pb_emptybot_gen.py gen | bake | read
"""
import os, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_emptybot")
DOCX = os.path.join(OUT, "emptybot.docx")
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
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')

NLINES = 30
BOTTOMS = list(range(300, 620, 10))   # twips


def txt(s):
    return '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % s


def sect(bottom, last):
    s = ('<w:sectPr>%s<w:pgSz w:w="11906" w:h="16838"/>'
         '<w:pgMar w:top="720" w:right="720" w:bottom="%d" w:left="720" '
         'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
         % ("" if last else '<w:type w:val="nextPage"/>', bottom))
    return s if last else '<w:p><w:pPr>%s</w:pPr></w:p>' % s


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for i, b in enumerate(BOTTOMS):
        tag = "%02d" % i
        for k in range(NLINES):
            body.append(txt("L%02d_%s" % (k, tag)))
        body.append("<w:p/>")                       # the empty under test
        body.append(txt("TGT_%s" % tag))
        body.append(sect(b, i == len(BOTTOMS) - 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           '><w:body>' + "".join(body) + "</w:body></w:document>")
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(BOTTOMS), "arms")


def read():
    import win32com.client as w
    app = w.DispatchEx("Word.Application"); app.Visible = False
    d = app.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    try:
        d.Repaginate()
        rows = {}
        n = d.Paragraphs.Count
        prev_tag = None
        for i in range(1, n + 1):
            r = d.Paragraphs(i).Range
            c = d.Range(r.Start, r.Start)
            t = r.Text.strip()
            if t.startswith("L29_"):
                prev_tag = t[4:]
                rows.setdefault(prev_tag, {})["L29"] = (c.Information(3), c.Information(6))
            elif t == "" and prev_tag is not None and "EMP" not in rows.get(prev_tag, {}):
                rows[prev_tag]["EMP"] = (c.Information(3), c.Information(6))
            elif t.startswith("TGT_"):
                tg = t[4:]
                rows.setdefault(tg, {})["TGT"] = (c.Information(3), c.Information(6))
                prev_tag = None
    finally:
        d.Close(False); app.Quit()
    print("%-4s %8s %9s %8s %8s %8s   %s" %
          ("arm", "bottom", "cbot", "L29_y", "EMP_y", "EMPpg-L29pg", "verdict"))
    for i, b in enumerate(BOTTOMS):
        tag = "%02d" % i
        r = rows.get(tag)
        if not r or "EMP" not in r or "L29" not in r:
            print("%-4s MISSING %s" % (tag, r)); continue
        cbot = 841.9 - b / 20.0
        same = r["EMP"][0] - r["L29"][0]
        print("%-4s %8.2f %9.2f %8.2f %8.2f %8d   %s" %
              (tag, b / 20.0, cbot, r["L29"][1], r["EMP"][1], same,
               "KEEP" if same == 0 else "PUSH"))


{"gen": gen, "read": read}[sys.argv[1]]()
