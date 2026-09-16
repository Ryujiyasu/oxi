# -*- coding: utf-8 -*-
"""What does Word do with a manual page break (`w:br w:type="page"`) inside a table cell?

technical__00c13e6a (compat 15): the Table 2-3 title row (tblHeader + cantSplit) holds a
paragraph that starts with THREE page-break runs before its text. Word draws the title as
the first line of the cell (row top 75.75, text 79.5) — the breaks leave no line and cause
no break. Oxi emits an empty line per extra break (two lines above the title, +21pt, and
the table's last row on that page falls to the next).

Sheet: 'before' body paragraph, a 2-row × 1-cell table (row 1 = the probe cell, row 2 =
'CELL B'), 'after' body paragraph. Arms: compat 14/15 × tblHeader 0/1 × cantSplit 0/1 ×
break position (start3 = three breaks then text; mid = 'X' + break + 'Y'; end = text then
break) + a body control (start3 in a body paragraph). Readout: per-line (page, y, text)
of the probe paragraph via per-character Information(3/6), plus page/y of before / CELL B /
after.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/cellbr'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
RPR = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr>'
BR = f'<w:r>{RPR}<w:br w:type="page"/></w:r>'


def settings(compat):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')


def run(t):
    return f'<w:r>{RPR}<w:t xml:space="preserve">{t}</w:t></w:r>'


def probe_para(pos):
    if pos == 'start3':
        body = BR + BR + BR + run('CELL A title')
    elif pos == 'mid':
        body = run('X part') + BR + run('Y part')
    else:
        body = run('CELL A title') + BR
    return f'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/>{RPR}</w:pPr>{body}</w:p>'


def document(hdr, cant, pos, where):
    if where == 'body':
        body = f'<w:p><w:r>{RPR}<w:t>before</w:t></w:r></w:p>{probe_para(pos)}<w:p><w:r>{RPR}<w:t>after</w:t></w:r></w:p>'
    else:
        trpr = '<w:trPr>' + ('<w:cantSplit/>' if cant else '') + ('<w:tblHeader/>' if hdr else '') + '</w:trPr>'
        tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
               f'<w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr>{trpr}<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>{probe_para(pos)}</w:tc></w:tr>'
               f'<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr><w:p><w:r>{RPR}<w:t>CELL B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
        body = f'<w:p><w:r>{RPR}<w:t>before</w:t></w:r></w:p>{tbl}<w:p><w:r>{RPR}<w:t>after</w:t></w:r></w:p>'
    sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


def lines_of(d, rng, trim_end):
    lines = []; last = None; cur = ''
    for k in range(rng.Start, rng.End - trim_end):
        r = d.Range(k, k)
        key = (r.Information(3), round(r.Information(6), 2))
        if key != last:
            if last is not None:
                lines.append((last, cur))
            cur = ''
        last = key; cur += d.Range(k, k + 1).Text
    if last is not None:
        lines.append((last, cur))
    return [(pg, y, t.replace('\r', '\\r').replace('\x0c', '\\f').replace('\x07', '')) for (pg, y), t in lines]


def pos_of(d, rng):
    r = d.Range(rng.Start, rng.Start)
    return (r.Information(3), round(r.Information(6), 2))


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = [(c, 0, 0, 'start3', 'body') for c in (14, 15)]
    for c in (14, 15):
        for hdr in (0, 1):
            for cant in (0, 1):
                for pos in ('start3', 'mid', 'end'):
                    arms.append((c, hdr, cant, pos, 'cell'))
    for compat, hdr, cant, pos, where in arms:
        at = OUT / f'c{compat}_h{hdr}_c{cant}_{pos}_{where}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', settings(compat)); z.writestr('word/document.xml', document(hdr, cant, pos, where))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            before = pos_of(d, d.Paragraphs(1).Range)
            if where == 'body':
                pr = d.Paragraphs(2).Range; trim = 0
                cellb = None
                after = pos_of(d, d.Paragraphs(3).Range)
            else:
                pr = d.Tables(1).Cell(1, 1).Range.Paragraphs(1).Range; trim = 1
                cellb = pos_of(d, d.Tables(1).Cell(2, 1).Range)
                after = pos_of(d, d.Paragraphs(d.Paragraphs.Count).Range)
            print(f'compat{compat} hdr{hdr} cant{cant} {pos:6} {where:4} before={before} probe={lines_of(d, pr, trim)} cellB={cellb} after={after} pages={d.ComputeStatistics(2)}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
