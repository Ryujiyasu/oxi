# -*- coding: utf-8 -*-
"""Is `<w:trHeight w:val="866" w:hRule="auto"/>` a minimum row height, or ignored?

en/reports__001f1397 row 22 carries trHeight 866 hRule="auto" and splits across pages 1-2;
on page 2 Word gives its remainder («FY», one line) one line, Oxi ~43pt = 866tw.

Arms: one-line rows with trHeight 866 under hRule auto / atLeast / exact / (absent rule),
each followed by a probe row; the distance between the two rows' text baselines = the
first row's height.  A second table splits an auto row across a page (filler before it).

    python tools/metrics/_pb_trheight_auto_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_trheight_auto")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
RPR = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr>'
PPR = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
RULES = ["auto", "atLeast", "exact", None]


def p(text):
    return '<w:p>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (PPR, RPR, text)


def row(cells, rule, val=866):
    tr = "" if rule == "none" else ('<w:trPr><w:trHeight w:val="%d"%s/></w:trPr>' % (val, (' w:hRule="%s"' % rule) if rule else ""))
    return "<w:tr>%s%s</w:tr>" % (tr, "".join('<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>%s</w:tc>' % c for c in cells))


def tbl(rows):
    b = "".join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="000000"/>' % x for x in ("top", "left", "bottom", "right", "insideH", "insideV"))
    return ('<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders>%s</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>%s</w:tbl>' % (b, "".join(rows)))


def build(split_fill):
    body = ""
    for k, r in enumerate(RULES):
        tag = r or "norule"
        body += tbl([row([p("H%s top" % tag), p("x")], r), row([p("H%s next" % tag), p("y")], "none")]) + p("")
    if split_fill:
        # a 3-line cell in an auto row that splits across the page; then a probe row
        body += "".join(p("fill %d" % i) for i in range(split_fill))
        body += tbl([row([p("S1 a") + p("S2 a") + p("S3 a"), p("short")], "auto"), row([p("S after"), p("z")], "none")])
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>'
           '</w:body></w:document>' % (W, body))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "trh_fill%d.docx" % split_fill)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/document.xml", doc)
    return path


def ys(pdf):
    import fitz
    out = {}
    for pi, pg in enumerate(fitz.open(pdf)):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(c["c"] for s in l["spans"] for c in s["chars"])
                o = l["spans"][0]["chars"][0]["origin"][1] if l["spans"] and l["spans"][0]["chars"] else None
                out.setdefault(t.strip(), (pi + 1, round(o, 2)))
    return out


if __name__ == "__main__":
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for fill in [int(x) for x in os.environ.get("PB_FILLS", "0,40,41,42").split(",")]:
            docx = build(fill); pdf = docx[:-5] + ".pdf"
            d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
            y = ys(pdf)
            if fill == 0:
                for r in RULES:
                    tag = r or "norule"
                    a, b = y.get("H%s top" % tag), y.get("H%s next" % tag)
                    print("rule %-8s row height %.2f" % (tag, b[1] - a[1]), flush=True)
            else:
                print("fill %d  %s" % (fill, {k: v for k, v in y.items() if k.startswith(("S", "fill %d" % (fill - 1)))}), flush=True)
    finally:
        w.Quit()
