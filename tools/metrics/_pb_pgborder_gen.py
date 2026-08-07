# -*- coding: utf-8 -*-
"""Does <w:pgBorders offsetFrom="page"> shrink the text area?

correspondence__000f9471 (A4, all margins 36pt, pgBorders offsetFrom=page
space=24 sz=18) puts its page-1 tail at a box bottom of 805.20 with a naive
content bottom of 805.92 - and Word still pushes it to page 2, so Word's real
content bottom is below 805.20 while the deepest ink on another page reaches
785.0.  Window: cbot in [785.0, 805.20).  The only structure that can explain
it is the page border.

Each arm is its own SECTION with its own pgBorders; a uniform 12pt TNR filler
(line=240 auto, before/after 0) is emitted and the page-1 line COUNT gives the
content bottom to one line (13.8pt).  Reading the first baseline on the page
also pins whether the TOP is inset.

  python _pb_pgborder_gen.py gen | bake | read
"""
import os, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_pgborder")
DOCX = os.path.join(OUT, "pgborder.docx")
PDF = os.path.join(OUT, "pgborder.pdf")

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
RPR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>'
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>' + RPR + '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>')

# (tag, offsetFrom, space, sz)   sz is eighths of a point
ARMS = [("A", None, None, None),      # no border at all (control)
        ("B", "page", 24, 18),        # the target's own geometry
        ("C", "page", 12, 18),
        ("D", "page", 31, 18),
        ("E", "page", 24, 4),
        ("F", "text", 24, 18),
        ("G", "page", 24, 48)]
NLINE = 60


def sect(tag, off, space, sz, last):
    b = ""
    if off is not None:
        s = ('<w:pgBorders w:offsetFrom="%s">' % off +
             "".join('<w:%s w:val="single" w:sz="%d" w:space="%d" w:color="auto"/>' % (e, sz, space)
                     for e in ("top", "left", "bottom", "right")) +
             "</w:pgBorders>")
        b = s
    sp = ('<w:pgSz w:w="11906" w:h="16838"/>'
          '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
          'w:header="708" w:footer="708" w:gutter="0"/>' + b)
    return sp if last else "<w:p><w:pPr><w:sectPr>" + sp + "</w:sectPr></w:pPr></w:p>"


def build():
    body = []
    for i, (tag, off, space, sz) in enumerate(ARMS):
        for n in range(NLINE):
            body.append('<w:p><w:r><w:t xml:space="preserve">%s%02d wwwwwwwwww</w:t></w:r></w:p>' % (tag, n))
        last = (i == len(ARMS) - 1)
        if not last:
            body.append(sect(tag, off, space, sz, False))
        else:
            body.append("<w:sectPr>" + sect(tag, off, space, sz, True) + "</w:sectPr>")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
            '><w:body>' + "".join(body) + '</w:body></w:document>')


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
    doc = fitz.open(PDF)
    per = {}
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
            t = "".join(rows[y]).strip()
            per.setdefault(t[:3], (pi, y))
    print("%-4s %-9s %5s %4s | %8s %8s %8s %s" %
          ("arm", "offsetFrom", "space", "sz", "firstY", "lastY", "nline", "note"))
    for tag, off, space, sz in ARMS:
        ys = [(p, y, n) for n in range(NLINE)
              for (p, y) in [per.get("%s%02d" % (tag, n), (None, None))] if p is not None]
        if not ys:
            print("%-4s missing" % tag); continue
        p0 = ys[0][0]
        first = [y for p, y, n in ys if p == p0]
        print("%-4s %-9s %5s %4s | %8.2f %8.2f %8d  cbot>=%.2f" %
              (tag, str(off), str(space), str(sz), first[0], first[-1], len(first),
               first[-1] + 2.66))


if __name__ == "__main__":
    {"gen": gen, "bake": bake, "read": read}[sys.argv[1]]()
