# -*- coding: utf-8 -*-
"""At a row split, does EACH CELL's remaining content restart at the continuation
cell top, or is the whole overflow shifted by ONE global delta?

Oxi's Step-1 re-anchor (2026-04-22, + S570/S719b/S817/S998) computes a single
`min_overflow_text_y` over the WHOLE row and shifts every overflow text element
by one `adjust`.  Word appears to restart each cell independently -- observed on
uk_local_spending p46/p47, where col-2's continuation and col-3's 'Code' land on
the SAME first continuation line even though their pre-split y differ.

To separate the two models the cells must have DIFFERENT line phases, so each
arm gives the two cells different font sizes:

    cell A : NA lines at FSA pt      cell B : NB lines at FSB pt

The row is forced to split by making it taller than the page.  Then:

    global-shift model : the cell whose first overflow line sits LOWER before
                         the split starts LOWER on the continuation page
    per-cell model     : both cells start at the same continuation top

  python _pb_cellanchor_gen.py gen | bake | read
"""
import os, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.abspath(os.path.join(HERE, "..", "..", "pipeline_data", "_pb_cellanchor"))
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

# (arm, font size A, n lines A, font size B, n lines B)
ARMS = [
    ("EQ",   12, 60, 12, 60),   # control: same phase in both cells
    ("AB1",  12, 60,  8, 90),   # different phases
    ("AB2",  12, 60, 16, 45),
    ("AB3",  16, 45,  8, 90),
]

def para(text, fs, tag):
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="%d"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="%d"/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (fs * 2, fs * 2, text))

def cell(w, lines, fs, tag):
    body = "".join(para("%s%03d" % (tag, i), fs, tag) for i in range(lines))
    return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc>' % (w, body))

def doc_xml():
    parts = []
    for (name, fsa, na, fsb, nb) in ARMS:
        row = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>'
               '<w:tblBorders>'
               '<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
               '<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
               '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
               '<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
               '<w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
               '<w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
               '</w:tblBorders></w:tblPr>'
               '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>'
               '<w:tr>%s%s</w:tr></w:tbl>'
               % (cell(4500, na, fsa, name + "A"), cell(4500, nb, fsb, name + "B")))
        marker = para("MARK-%s" % name, 12, "M")
        brk = ('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        parts.append(marker + row + brk)
    body = "".join(parts)
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
            ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document %s><w:body>%s%s</w:body></w:document>' % (NS, body, sect))

def gen():
    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, "cellanchor.docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc_xml())
    print("built", p, "(%d arms)" % len(ARMS))

def read():
    import fitz
    pdf = os.path.join(OUT, "cellanchor.pdf")
    d = fitz.open(pdf)
    print("pages", d.page_count)
    for pi in range(d.page_count):
        spans = []
        for b in d[pi].get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l.get("spans", []):
                    t = s["text"].strip()
                    if t:
                        spans.append((round(s["origin"][1], 2), round(s["origin"][0], 2), t))
        if not spans:
            continue
        spans.sort()
        # first two spans per column bucket (x < 300 = cell A, else cell B)
        a = [s for s in spans if s[1] < 300][:2]
        b = [s for s in spans if s[1] >= 300][:2]
        print("p%-2d  A:%-34s B:%s"
              % (pi + 1,
                 " ".join("%s@%.2f" % (t, y) for y, x, t in a),
                 " ".join("%s@%.2f" % (t, y) for y, x, t in b)))

if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    {"gen": gen, "read": read}[cmd]()
