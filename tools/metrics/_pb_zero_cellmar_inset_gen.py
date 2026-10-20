# -*- coding: utf-8 -*-
"""Where does cell text start when the cell margins are zero?

golden probethdr / probezgridspan (styles part without a default table style, so the
cell margins are 0): Word's cell text starts 0.72pt right of the cell edge (border sz 4
drawn 70.94..71.42, text 71.66); Oxi starts on the edge and wraps one character early.

Arms: left/right border sz none / 4 / 8 / 12 / 24 x cell margin 0 / 20tw (explicit
tblCellMar). Per arm: first glyph x - cell edge, and the right-most glyph's right edge
against the right cell edge on a line of ideographs long enough to fill the cell.

    python tools/metrics/_pb_zero_cellmar_inset_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_zero_cellmar_inset")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>' % W)
TEXT = "本項に定める事項については関係法令及び本規程の趣旨に照らし善良なる管理者の注意をもって誠実かつ適切に取り扱うもの"


def build(sz, mar):
    b = "".join(('<w:%s w:val="nil"/>' % x) if sz == 0 else ('<w:%s w:val="single" w:sz="%d" w:space="0" w:color="auto"/>' % (x, sz))
                for x in ("top", "left", "bottom", "right"))
    cm = '' if mar is None else ('<w:tblCellMar><w:left w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/></w:tblCellMar>' % (mar, mar))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="7000" w:type="dxa"/><w:tblBorders>%s</w:tblBorders>%s</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="7000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr>'
           '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>%s</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' % (b, cm, TEXT))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s<w:p/>'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/>'
           '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>' % (W, tbl))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "zc_b%d_m%s.docx" % (sz, mar))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    return path


def first_line(pdf):
    import fitz
    pg = fitz.open(pdf)[0]
    chars = [c for b in pg.get_text("rawdict")["blocks"] for l in b.get("lines", []) for s in l["spans"] for c in s["chars"] if c["c"].strip()]
    y0 = min(c["origin"][1] for c in chars)
    row = sorted([c for c in chars if abs(c["origin"][1] - y0) < 1], key=lambda c: c["origin"][0])
    return row[0]["origin"][0], row[-1]["bbox"][2], len(row)


if __name__ == "__main__":
    import win32com.client
    edge_l, edge_r = 1418 / 20, 1418 / 20 + 350
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for mar in (None, 0, 20):
            for sz in (0, 4, 8, 12, 24):
                docx = build(sz, mar); pdf = docx[:-5] + ".pdf"
                d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
                x0, x1, n = first_line(pdf)
                print("mar %-4s border sz %2d: text left %.2f (edge +%.2f)  right %.2f (edge %+.2f)  chars %d"
                      % (mar, sz, x0, x0 - edge_l, x1, x1 - edge_r, n), flush=True)
    finally:
        w.Quit()
