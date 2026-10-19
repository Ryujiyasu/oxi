# -*- coding: utf-8 -*-
"""Where does Word break a vMerge group whose restart cell is taller than its first row?

en/reports__001f1397 (no sectPr): row 60 = «2021/22» | «FY» | vMerge restart «3% plus an
additional / 1% non-consolidated.» (2 lines), row 61 = vMerge | «FY top scale» | vMerge.
Word keeps row 60 on page 2 with the merged cell's first line and carries «1% non-
consolidated.» to page 3 beside «FY top scale»; Oxi moves both rows to page 3.

Arms: filler paragraphs push the group to the page bottom; the bottom margin is swept so
that 0 / 1 / 2 / 3 lines fit after the filler.  Per arm: the page of each cell text.

    python tools/metrics/_pb_vmerge_pagebreak_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_vmerge_pagebreak")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
RPR = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr>'
PPR = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
N_FILL = int(os.environ.get("PB_FILL", "40"))
MERGED = os.environ.get("PB_MERGED", "2")   # paragraphs in the restart cell


def p(text):
    return '<w:p>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (PPR, RPR, text)


def tc(w, body, vm=None):
    v = '<w:vMerge w:val="restart"/>' if vm == "restart" else ('<w:vMerge/>' if vm == "cont" else "")
    return '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s</w:tcPr>%s</w:tc>' % (w, v, body)


def table():
    merged = "".join(p("M%d merged line" % (i + 1)) for i in range(int(MERGED)))
    rows = [
        tc(2000, p("R0 first")) + tc(2000, p("R0 second")) + tc(3000, p("R0 third")),
        tc(2000, p("A1 group"), "restart") + tc(2000, p("A2 first row")) + tc(3000, merged, "restart"),
        tc(2000, p(""), "cont") + tc(2000, p("B2 second row")) + tc(3000, p(""), "cont"),
        tc(2000, p("C1 after")) + tc(2000, p("C2 after")) + tc(3000, p("C3 after")),
    ]
    tr = "".join('<w:tr>%s</w:tr>' % r for r in rows)
    borders = "".join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="000000"/>' % b for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    return ('<w:tbl><w:tblPr><w:tblW w:w="7000" w:type="dxa"/><w:tblBorders>%s</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>%s</w:tbl>' % (borders, tr))


def build(bottom):
    body = "".join(p("filler %d" % i) for i in range(N_FILL)) + table() + p("tail")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="%d" w:left="1440" w:right="1440"/></w:sectPr>'
           '</w:body></w:document>' % (W, body, bottom))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "vm_m%s_b%d.docx" % (MERGED, bottom))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    import fitz, win32com.client
    bottoms = [int(b) for b in os.environ.get("PB_BOTTOMS", "").split(",") if b] or list(range(1440, 1440 + 20 * 26, 20))
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        prev = None
        for b in bottoms:
            docx = build(b); pdf = docx[:-5] + ".pdf"
            d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
            where = {}
            for pi, pg in enumerate(fitz.open(pdf)):
                txt = pg.get_text()
                for key in ("R0 first", "A1 group", "A2 first row", "M1 merged", "M2 merged", "M3 merged", "B2 second row", "C1 after"):
                    if key in txt and key not in where:
                        where[key] = pi + 1
            row = " ".join("%s=%s" % (k.split()[0], v) for k, v in where.items())
            if row != prev:
                print("bottom %4d  %s" % (b, row), flush=True)
                prev = row
    finally:
        w.Quit()
