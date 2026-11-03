# -*- coding: utf-8 -*-
"""Does a justified pure-Latin line in a JAPANESE document shrink its spaces?

ja/forms__00e53925 (compat 15, compressPunctuation, Century 10.5 in a 3266tw cell):
Word keeps «Connections Luxury in Madeira» (hmtx 154.57) in a 152.5 cell, its
three spaces drawn 2.16 instead of 2.92 (the S825 quarter). Oxi's S799 / S1082
space shrink is gated to documents with no real CJK text, so the cell wrapped.

Arms: jc (both / left) x context (body / cell) x doc (cjk: a Japanese paragraph
above / latin: none). Each document holds 41 copies at inner widths stepping 2tw
around the natural width; readout = the narrowest width keeping «Madeira» on line 1.
"""
import os, sys, zipfile, itertools
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/cjkdoc_latin_shrink'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
RPR = ('<w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:sz w:val="21"/>'
       '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr>')
TEXT = 'Connections Luxury in Madeira'
NAT_TW = 3091  # hmtx 154.567pt


def para(text, jc, ind_r=0):
    ind = f'<w:ind w:right="{ind_r}"/>' if ind_r else ''
    return f'<w:p><w:pPr><w:spacing w:after="0"/>{ind}<w:jc w:val="{jc}"/>{RPR}</w:pPr><w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def document(jc, ctx, doc):
    body = [para('日本語の段落です。', 'left')] if doc == 'cjk' else [para('An English paragraph.', 'left')]
    for k in range(-80, 9, 2):
        inner = NAT_TW + k
        body.append(para(f'K{k:+d}', 'left'))
        if ctx == 'cell':
            body.append(f'<w:tbl><w:tblPr><w:tblW w:w="{inner + 216}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
                        '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
                        f'<w:tblGrid><w:gridCol w:w="{inner + 216}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="{inner + 216}" w:type="dxa"/></w:tcPr>'
                        f'{para(TEXT, jc)}</w:tc></w:tr></w:tbl>')
        else:
            body.append(para(TEXT, jc, ind_r=9638 - inner))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1134" w:bottom="720" w:left="1134" '
            'w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>')


def readout(pdf):
    import pymupdf
    res = {}; label = None
    for P in pymupdf.open(pdf):
        rows = {}
        for b in P.get_text('rawdict')['blocks']:
            for l in b.get('lines', []):
                for s in l['spans']:
                    for c in s['chars']:
                        rows.setdefault(round(c['bbox'][1], 0), []).append((c['bbox'][0], c['c']))
        for y in sorted(rows):
            r = ''.join(c for _, c in sorted(rows[y]))
            if r.startswith('K') and r[1:2] in '+-':
                label = int(r[1:].strip()); res[label] = []
            elif label is not None:
                res[label].append(r)
    keep = sorted(k for k, r in res.items() if r and 'Madeira' in r[0])
    return keep


if __name__ == '__main__':
    import win32com.client
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for jc, ctx, doc in itertools.product(('both', 'left'), ('cell', 'body'), ('cjk', 'latin')):
            name = f'{jc}_{ctx}_{doc}'
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', SETTINGS)
                z.writestr('word/document.xml', document(jc, ctx, doc))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            keep = readout(pdf)
            mk = min(keep) if keep else None
            print(f'{name:20} min_keep {mk}tw  allow {None if mk is None else -mk / 20}pt', flush=True)
    finally:
        app.Quit()
