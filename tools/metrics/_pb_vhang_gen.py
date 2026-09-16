# -*- coding: utf-8 -*-
"""Does Word hang a column-final 。 past the column end in vertical text, or push it (with the
character before it) to the next column?

creative__25b9ec89 (compat 14, ＭＳ Ｐ明朝 11pt): 'ふと、…気がした。' (43 chars, 426.5pt of a
425.2pt column) breaks 41 + 'た。' in Word; Oxi (vertical advance checkpoint) hangs the 。 and
packs all 43.

Sheet: vertical section (column 425.2pt), ＭＳ 明朝 12pt (35.43 em per column). Arms: text of
35 chars + 。 + tail (the 。 is exactly the 36th = past the end), 34 chars + 。 + tail (。 is the
35th = last fitting char), 34 chars + 、 + tail, compat 14 / 15. Readout: characters in column 1
(distinct Information(5) x of the paragraph's characters).
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/vhang'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
KANA = 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん'


def settings(compat):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')


FONT = os.environ.get('FONT', 'ＭＳ 明朝')
SZ = os.environ.get('SZ', '24')


def document(text):
    rpr = f'<w:rPr><w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}" w:hint="eastAsia"/><w:kern w:val="0"/><w:sz w:val="{SZ}"/></w:rPr>'
    body = f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>'
    sect = ('<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1701" w:right="1985" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:cols w:space="425"/><w:textDirection w:val="tbRl"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = {f'n{n}_maru': KANA[:n] + '。' + KANA[:6] for n in (int(x) for x in os.environ['NS'].split(','))} if os.environ.get('NS') else {
        'n35_maru': KANA[:35] + '。' + KANA[:6],
        'n34_maru': KANA[:34] + '。' + KANA[:6],
        'n34_ten': KANA[:34] + '、' + KANA[:6],
        'n35_ten': KANA[:35] + '、' + KANA[:6],
        'n34_kagi': KANA[:34] + '」' + KANA[:6],
        'n35_plain': KANA[:35] + KANA[:6],
    }
    for compat in (14, 15):
        for name, text in arms.items():
            at = OUT / f'c{compat}_{name}_{FONT.replace(" ", "")}_{SZ}.docx'
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
                z.writestr('word/settings.xml', settings(compat)); z.writestr('word/document.xml', document(text))
            d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
            try:
                p = d.Paragraphs(1).Range
                cols = []; last = None; cnt = 0; chars = ''
                for c in range(p.Start, p.End - 1):
                    x = round(d.Range(c, c).Information(5), 1)
                    if x != last:
                        if last is not None: cols.append((cnt, chars[-3:]))
                        last = x; cnt = 0; chars = ''
                    cnt += 1; chars += d.Range(c, c + 1).Text
                if last is not None: cols.append((cnt, chars[-3:]))
                print(f'compat{compat} {name:10} cols(n, last3)={cols}', flush=True)
            finally:
                d.Close(False)
finally:
    app.Quit()
