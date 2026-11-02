# -*- coding: utf-8 -*-
"""Latin advances in a run whose (missing) font Word replaces with Microsoft
YaHei, and whether «Q.　「生きてみる」ボタンを押すとどうなりますか？» fits one line.

ja/reference__2296511c (compat 14, compressPunctuation, balanceSingleByteDoubleByteWidth,
run fonts all «Noto Sans Japanese Bold» (not installed) + hint eastAsia, 10.5pt,
3 columns of 4849tw = 242.45): Word draws Q 7.92 / . 2.52 / U+3000 5.76 and the
line ends at 242.45 exactly; YaHei's own Q is 0.815 em (8.56) and Oxi gives «Q.»
11.58, so Oxi wraps «か？».

Arms: font Noto-missing / Microsoft YaHei x hint eastAsia / none x balance on /
off, the paragraph alone in a 4849tw column. One Word process per arm (Word
caches missing-font substitutes per process).
Readout: Word PDF glyph origins of the line, the line count.
"""
import os, sys, zipfile, itertools, subprocess
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/yahei_latin'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
TEXT = 'Q.　「生きてみる」ボタンを押すとどうなりますか？'


def settings(bal):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat>' + ('<w:balanceSingleByteDoubleByteWidth/>' if bal else '') +
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat></w:settings>')


def document(face, hint):
    h = ' w:hint="eastAsia"' if hint else ''
    rpr = f'<w:rPr><w:rFonts{h} w:ascii="{face}" w:hAnsi="{face}" w:eastAsia="{face}" w:cs="{face}"/><w:sz w:val="21"/></w:rPr>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            f'<w:p><w:pPr><w:spacing w:after="0" w:line="0" w:lineRule="atLeast"/></w:pPr><w:r>{rpr}<w:t>{TEXT}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="{11906 - 1440 - 4849}" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


ONE = r'''
import sys, win32com.client, pymupdf
sys.stdout.reconfigure(encoding='utf-8')
at = sys.argv[1]; pdf = at[:-5] + '.pdf'
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
try:
    d = app.Documents.Open(at, ReadOnly=True)
    try: d.ExportAsFixedFormat(pdf, 17)
    finally: d.Close(False)
finally:
    app.Quit()
P = pymupdf.open(pdf)[0]
lines = []
for b in P.get_text('rawdict')['blocks']:
    for l in b.get('lines', []):
        cs = [c for s in l['spans'] for c in s['chars']]
        lines.append((round(l['bbox'][1], 1), [(c['c'], round(c['origin'][0], 2)) for c in cs], l['spans'][0]['font']))
lines.sort()
print(len(lines), lines[0][2], lines[0][1][:6], 'end', lines[0][1][-2:])
'''

if __name__ == '__main__':
    one = os.path.abspath(os.path.join(OUT, '_one.py'))
    open(one, 'w', encoding='utf-8').write(ONE)
    for face, hint, bal in itertools.product(('Noto Sans Japanese Bold', 'Microsoft YaHei'), (True, False), (True, False)):
        name = f'{"noto" if face.startswith("Noto") else "yahei"}_{"hint" if hint else "nohint"}_{"bal" if bal else "nobal"}'
        at = os.path.abspath(os.path.join(OUT, name + '.docx'))
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
            z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', settings(bal))
            z.writestr('word/document.xml', document(face, hint))
        r = subprocess.run([sys.executable, one, at], capture_output=True, text=True, encoding='utf-8')
        print(f'{name:22}', (r.stdout or r.stderr).strip()[-400:], flush=True)
