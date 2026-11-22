# -*- coding: utf-8 -*-
"""Does a fullwidth colon / semicolon before an opening bracket give up its half em, like «・「»?

blind-A ja/reference__0ae2d13d: Word sets «場合：「家族登録」» (HG丸ｺﾞｼｯｸM-PRO 12, compat 15,
compressPunctuation, jc left, slack line) with ： advancing 6.00; Oxi 12.00. The Word-PDF census
(`colon_open_scan`) finds ： before an opener at ~0.5 em in 15+ docs across ＭＳ 明朝 / ＭＳ ゴシック /
HG丸ｺﾞｼｯｸM-PRO / Meiryo, ~1.0 em before a kana or kanji.
Arms: font x (compressPunctuation | doNotCompress) x (jc left | both) x (no grid | linesAndChars),
texts «：「», «：（», «；「», «：あ», «・「». Readout: PDF glyph-origin advances of the first three chars.
"""
import os, sys, zipfile, itertools
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/pair_colon_bracket'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
FONTS = {'mincho': 'ＭＳ 明朝', 'maru': 'HG丸ｺﾞｼｯｸM-PRO'}
TEXTS = ['あ：「働きやすい」', 'あ：（働きやすい）', 'あ；「働きやすい」', 'あ：あ働きやすい', 'あ・「働きやすい」']


def rpr(font):
    return f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:hint="eastAsia"/><w:sz w:val="24"/><w:lang w:eastAsia="ja-JP"/></w:rPr>'


def settings(ctrl):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:characterSpacingControl w:val="{ctrl}"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def document(font, jc, grid):
    r = rpr(font)
    body = ''.join(f'<w:p><w:pPr><w:jc w:val="{jc}"/>{r}</w:pPr><w:r>{r}<w:t>{t}</w:t></w:r></w:p>' for t in TEXTS)
    g = ('<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="1060"/>' if grid
         else '<w:docGrid w:type="lines" w:linePitch="360"/>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            f'w:header="851" w:footer="992" w:gutter="0"/>{g}</w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for fk, ctrl, jc, grid in itertools.product(FONTS, ('compressPunctuation', 'doNotCompress'), ('left', 'both'), (False, True)):
            name = f'{fk}_{ctrl[:4]}_{jc}_{"grid" if grid else "nogrid"}'
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', settings(ctrl))
                z.writestr('word/document.xml', document(FONTS[fk], jc, grid))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = pymupdf.open(pdf)[0]
            for b in P.get_text('rawdict')['blocks']:
                for l in b.get('lines', []):
                    ch = [c for s in l['spans'] for c in s['chars']]
                    xs = [c['origin'][0] for c in ch]
                    adv = [round(xs[i + 1] - xs[i], 2) for i in range(min(3, len(xs) - 1))]
                    print(f'{name:32} {"".join(c["c"] for c in ch)[:9]:12} adv {adv}', flush=True)
    finally:
        app.Quit()
