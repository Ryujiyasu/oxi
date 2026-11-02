# -*- coding: utf-8 -*-
"""Where does Word put a ruby and its base in a VERTICAL (tbRl) section?

ja/reference__6358c5e9 (ＭＳ 明朝 12, linesAndChars linePitch 326, ruby hps 12 /
hpsRaise 22 / distributeSpace): a ruby column is two grid lines (32.6pt) wide;
the base em box sits 6.97pt in from the column's left edge (16.3 - 6 - 3.33),
the ruby column's centre 10.92 right of the base's centre. Oxi centred the base
in the 32.6pt (3.33 too far right).

Arms: base size 10.5 / 12 x ruby hps 6 / 8 x hpsRaise (hps-derived) 14 / 22 / 30
x grid linesAndChars 326 / none. Each document: a plain column, then a column
«夏草や南州無念» with ruby «なんしゅう» on «南州», then a plain column.
Paragraphs: after 0, line 240 auto (a styles-less package otherwise takes
Word's application default after 8 / 1.08).
Readout: Word PDF glyph boxes (x) of 夏, 南 and the ruby kana, and the plain
columns' x, so the column pitch and the base offset come out together.
"""
import os, sys, zipfile, itertools
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/vert_ruby'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
FONT = 'ＭＳ 明朝'


def rpr(sz):
    return f'<w:rPr><w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}" w:hint="eastAsia"/><w:sz w:val="{sz}"/></w:rPr>'


def document(base_hp, hps, raise_hp, grid):
    ruby = (f'<w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="{hps}"/><w:hpsRaise w:val="{raise_hp}"/>'
            f'<w:hpsBaseText w:val="{base_hp}"/><w:lid w:val="ja-JP"/></w:rubyPr><w:rt><w:r>{rpr(hps)}<w:t>なんしゅう</w:t></w:r></w:rt>'
            f'<w:rubyBase><w:r>{rpr(base_hp)}<w:t>南州</w:t></w:r></w:rubyBase></w:ruby></w:r>')
    ppr = '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    plain = lambda t: f'<w:p>{ppr}<w:r>{rpr(base_hp)}<w:t>{t}</w:t></w:r></w:p>'
    body = (plain('一行目の文') + f'<w:p>{ppr}<w:r>{rpr(base_hp)}<w:t>夏草や</w:t></w:r>{ruby}<w:r>{rpr(base_hp)}<w:t>無念</w:t></w:r></w:p>'
            + plain('三行目の文') + plain('四行目の文'))
    g = '<w:docGrid w:type="linesAndChars" w:linePitch="326"/>' if grid else '<w:docGrid w:linePitch="360"/>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1077" w:bottom="1440" w:left="1077" '
            f'w:header="851" w:footer="992" w:gutter="0"/><w:textDirection w:val="tbRl"/>{g}</w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for base_hp, hps, raise_hp, grid in itertools.product((21, 24), (12, 16), (14, 22, 30), (True, False)):
            name = f'b{base_hp}_r{hps}_h{raise_hp}_{"grid" if grid else "nogrid"}'
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', SETTINGS)
                z.writestr('word/document.xml', document(base_hp, hps, raise_hp, grid))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = pymupdf.open(pdf)[0]
            got = {}
            for b in P.get_text('rawdict')['blocks']:
                for l in b.get('lines', []):
                    for s in l['spans']:
                        for c in s['chars']:
                            key = c['c'] + ('r' if s['size'] < base_hp / 2 - 0.1 else '')
                            got.setdefault(key, (round(c['bbox'][0], 2), round(c['bbox'][1], 2)))
            pick = ['一', '夏', '南', '州', 'なr', 'うr', '三', '四']
            print(f'{name:22} ' + ' '.join(f'{k}{got.get(k)}' for k in pick), flush=True)
    finally:
        app.Quit()
