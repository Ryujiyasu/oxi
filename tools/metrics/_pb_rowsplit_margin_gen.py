# -*- coding: utf-8 -*-
"""How much room does Word need to start a split table row at the page bottom?

golden probeqtblstyle: a 2-cell row (cell margins top/bottom 60tw from the table style,
18pt grid lines) reaches the page bottom with 22.9pt left; Oxi starts it there (first
line 18 + top margin 3 = 21 fits), Word moves the whole row.

Arms: cell top/bottom margin 0 / 60 / 120 tw; filler paragraphs push the table so the
row starts a little above the bottom; the bottom page margin is swept in 0.4pt steps.
Per arm: the smallest room (page bottom - row top) at which Word still starts the row.

    python tools/metrics/_pb_rowsplit_margin_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_rowsplit_margin")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
TEXT = "本項に定める事項については、関係法令及び本規程の趣旨に照らし、善良なる管理者の注意をもって、誠実かつ適切に取り扱うものとし、これに関連して生じた一切の事項は、別に定めるところにより処理しなければならない。"
PPR = '<w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>'


def p(t):
    return '<w:p>%s<w:r><w:t>%s</w:t></w:r></w:p>' % (PPR, t)


def build(mar, bottom, fill):
    cm = '<w:tblCellMar><w:top w:w="%d" w:type="dxa"/><w:bottom w:w="%d" w:type="dxa"/></w:tblCellMar>' % (mar, mar)
    b = "".join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="auto"/>' % x for x in ("top", "left", "bottom", "right", "insideH", "insideV"))
    row = ('<w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>%s</w:tc>'
           '<w:tc><w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr>' % (p("ROWA"), p("MARKA" + TEXT)))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="8500" w:type="dxa"/><w:tblBorders>%s</w:tblBorders>%s</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="7000"/></w:tblGrid>%s</w:tbl>' % (b, cm, row))
    body = "".join(p("filler %d 行" % i) for i in range(fill)) + tbl + p("tail")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="%d" w:left="1418"/>'
           '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>' % (W, body, bottom))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "rs_m%d_b%d.docx" % (mar, bottom))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/document.xml", doc)
    return path


def row_start_page(pdf, fill):
    """1 = the row starts on the page of the last filler, 2 = it was pushed."""
    import fitz
    pf = pm = None
    for pi, pg in enumerate(fitz.open(pdf)):
        t = pg.get_text()
        if "filler %d " % (fill - 1) in t:
            pf = pi
        if "MARKA" in t and pm is None:
            pm = pi
    return None if pf is None or pm is None else (1 if pm == pf else 2)


if __name__ == "__main__":
    import win32com.client
    fill = int(os.environ.get("PB_FILL", "18"))
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for mar in (0, 60, 120):
            res = []
            for bottom in range(1418, 1418 + 20 * 30, 8):
                docx = build(mar, bottom, fill); pdf = docx[:-5] + ".pdf"
                d = w.Documents.Open(docx, ReadOnly=True)
                # page bottom and the row top straight from Word's layout
                # where the row WOULD start on page 1: below the last filler
                y_last = d.Paragraphs(fill).Range.Information(6)
                y_prev = d.Paragraphs(fill - 1).Range.Information(6)
                top = y_last + (y_last - y_prev)
                d.ExportAsFixedFormat(pdf, 17); d.Close(False)
                room = (16838 - bottom) / 20 - top
                res.append((round(room, 2), row_start_page(pdf, fill)))
            starts = [r for r, pgn in res if pgn == 1]
            pushes = [r for r, pgn in res if pgn == 2]
            print("cell margin %3d tw: min room where the row starts on p1 %s, max room where it is pushed %s"
                  % (mar, min(starts) if starts else None, max(pushes) if pushes else None), flush=True)
    finally:
        w.Quit()
