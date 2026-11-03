# -*- coding: utf-8 -*-
"""Does a justified TABLE CELL squeeze its CJK<->digit autoSpace gap to keep its last unit?

ja/reports__1fd371a8 (compat 15, docGrid lines, doNotCompress, jc=both, BIZ UDPゴシック
10.5 in a 1134tw cell): Word sets «317万人» on one line with the 7|万 gap at 2.30pt
(the other rows' gaps are 2.66), Oxi wraps «万人». S1585 derived the squeeze for BODY
lines only (s476_body).

One document per arm, each holding 41 one-cell tables whose inner width steps by 2tw
around the natural width. Readout: the narrowest inner width that keeps the last
character on the first line (Word PDF), i.e. allow = natural - that width.
Arms: text x jc x context x characterSpacingControl.
"""
import os, sys, zipfile, itertools
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/cell_gap_squeeze'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
FONT = 'BIZ UDPゴシック'
RPR = (f'<w:rPr><w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/>'
       '<w:sz w:val="21"/><w:lang w:eastAsia="ja-JP"/></w:rPr>')
TEXTS = {'g1': '317万人', 'g2': '万317万人', 'g1x': '1499万人',
         # mark + gap: does the gap credit add to the 、 capacity (compressPunctuation)?
         'm0': 'あ、いう万人', 'm1': 'あ、い317万人'}
# natural widths (pt) from the face's hmtx at 10.5pt; the sweep spans -4.0..+0.4
ADV = {'3': 1556, '1': 1290, '7': 1556, '4': 1556, '9': 1556}


def natural(text):
    em = sum(ADV.get(c, 2048) for c in text) / 2048 * 10.5
    gaps = sum(1 for a, b in zip(text, text[1:]) if a.isdigit() != b.isdigit())
    return em + gaps * 10.5 / 4


def settings(ctrl):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            f'<w:characterSpacingControl w:val="{ctrl}"/><w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def para(text, jc, ind_r=0):
    ind = f'<w:ind w:right="{ind_r}"/>' if ind_r else ''
    return (f'<w:p><w:pPr><w:spacing w:after="0"/>{ind}<w:jc w:val="{jc}"/>{RPR}</w:pPr>'
            f'<w:r>{RPR}<w:t>{text}</w:t></w:r></w:p>')


def document(text, jc, ctx):
    nat_tw = round(natural(text) * 20)
    body = []
    for k in (range(-240, 9, 4) if text.startswith("あ") else range(-80, 9, 2)):
        inner = nat_tw + k
        body.append(para(f'K{k:+d}', 'left'))
        if ctx == 'cell':
            body.append(f'<w:tbl><w:tblPr><w:tblW w:w="{inner + 216}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
                        '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>'
                        '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/></w:tblBorders>'
                        '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
                        f'<w:tblGrid><w:gridCol w:w="{inner + 216}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="{inner + 216}" w:type="dxa"/></w:tcPr>'
                        f'{para(text, jc)}</w:tc></w:tr></w:tbl>')
        else:
            body.append(para(text, jc, ind_r=9638 - inner))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1134" w:bottom="720" w:left="1134" '
            'w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>')


def readout(pdf, text):
    import pymupdf
    first, last = text[0], text[-1]
    res = {}
    for P in pymupdf.open(pdf):
        lines = []
        for b in P.get_text('rawdict')['blocks']:
            for l in b.get('lines', []):
                for s in l['spans']:
                    for c in s['chars']:
                        lines.append((round(c['bbox'][1], 1), c['bbox'][0], c['c']))
        lines.sort()
        # group by label K..
        label = None; rows = []
        cur = []
        ys = sorted({y for y, _, _ in lines})
        for y in ys:
            row = ''.join(c for yy, x, c in sorted((t for t in lines if t[0] == y), key=lambda t: t[1]))
            rows.append(row)
        for row in rows:
            if row.startswith('K'):
                label = int(row[1:]); res[label] = []
            elif label is not None:
                res[label].append(row)
    keep = sorted(k for k, r in res.items() if r and r[0].replace(' ', '') == text)
    wrap = sorted(k for k, r in res.items() if r and r[0].replace(' ', '') != text)
    return keep, wrap


if __name__ == '__main__':
    import win32com.client
    arms = list(itertools.product(TEXTS, ('both', 'left'), ('cell', 'body'), ('doNotCompress', 'compressPunctuation')))
    if len(sys.argv) > 1:
        arms = [a for a in arms if all(t in a for t in sys.argv[1].split(','))]
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for t, jc, ctx, ctrl in arms:
            name = f'{t}_{jc}_{ctx}_{ctrl}'
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', settings(ctrl))
                z.writestr('word/document.xml', document(TEXTS[t], jc, ctx))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            keep, wrap = readout(pdf, TEXTS[t])
            flip = min(keep) if keep else None
            allow = None if flip is None else -flip / 20
            print(f'{name:42} natural {natural(TEXTS[t]):6.2f}  min_keep {flip}tw  allow {allow}pt  (wrap max {max(wrap) if wrap else None})', flush=True)
    finally:
        app.Quit()
