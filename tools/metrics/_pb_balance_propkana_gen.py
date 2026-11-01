# -*- coding: utf-8 -*-
"""Does balanceSingleByteDoubleByteWidth double a run's w:spacing on PROPORTIONAL
kana (a P face), or only on full-width characters?

ja/correspondence__17884aa3 (HGP創英角ﾎﾟｯﾌﾟ体 10.5, w:spacing 22, linesAndChars
charSpace 9011, balance): Word advances kanji 14.87 (12.7 grid + 2 x 1.1) and
the kana に/も 11.75; Oxi gives the kana 12.9 -- 1.15 more, one spacing.

Arms: face HGP創英角ﾎﾟｯﾌﾟ体 / ＭＳ Ｐゴシック / ＭＳ ゴシック x spacing 0 / 22 / 44 x
balance on / off x grid linesAndChars (charSpace 9011) / none. Text 出席にも欠席に.
Readout: Word PDF per-character advances.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/balance_propkana'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
FACES = {'pop': 'HGP創英角ﾎﾟｯﾌﾟ体', 'pgo': 'ＭＳ Ｐゴシック', 'go': 'ＭＳ ゴシック'}
TEXT = '出席にも欠席に'


def settings(bal):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat>' + ('<w:balanceSingleByteDoubleByteWidth/>' if bal else '') +
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def document(face, sp, grid):
    rpr = f'<w:rPr><w:rFonts w:ascii="{face}" w:eastAsia="{face}" w:hAnsi="{face}" w:hint="eastAsia"/><w:spacing w:val="{sp}"/><w:sz w:val="21"/></w:rPr>'
    g = '<w:docGrid w:type="linesAndChars" w:linePitch="320" w:charSpace="9011"/>' if grid else '<w:docGrid w:linePitch="360"/>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body><w:p><w:r>{rpr}<w:t>{TEXT}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            f'w:header="720" w:footer="720" w:gutter="0"/>{g}</w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for fk, face in FACES.items():
            for grid in (True, False):
                for bal in (True, False):
                    for sp in (0, 22, 44):
                        name = f'{fk}_{"grid" if grid else "nogrid"}_{"bal" if bal else "nobal"}_sp{sp}'
                        at = os.path.abspath(os.path.join(OUT, name + '.docx'))
                        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                            z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', settings(bal))
                            z.writestr('word/document.xml', document(face, sp, grid))
                        pdf = at[:-5] + '.pdf'
                        d = app.Documents.Open(at, ReadOnly=True)
                        try:
                            d.ExportAsFixedFormat(pdf, 17)
                        finally:
                            d.Close(False)
                        P = pymupdf.open(pdf)[0]
                        xs = [(c['c'], c['origin'][0]) for b in P.get_text('rawdict')['blocks'] for l in b.get('lines', [])
                              for s in l['spans'] for c in s['chars'] if c['c'].strip()]
                        adv = [round(b2[1] - a[1], 2) for a, b2 in zip(xs, xs[1:])]
                        print(f'{name:22} ' + ' '.join(f'{a[0]}{v}' for a, v in zip(xs, adv)), flush=True)
    finally:
        app.Quit()
