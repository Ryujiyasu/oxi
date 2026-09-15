# -*- coding: utf-8 -*-
"""Greek / Cyrillic letters and the multiply / divide signs in a CJK document:
full-width or proportional, and does the eastAsia hint decide?

educational__0ad73366914e501f p2: 「同じ軌道についてσ,π結合を…」 (hint=eastAsia
on every run) is 4 lines in Word (43/41/41/3 chars) and 3 in Oxi (45/42/41),
which puts one paragraph more on the page. Oxi priced σ/π from the ascii face
(0.5em fallback); Word's line 0 only closes at 43 chars if σ and π are 1em.

Each arm: 6 paragraphs 「あ」+ 20 x ch +「あ」, ch in σ π α Я × ÷, in one
eastAsia face x one ascii face x hint on/off. Readout: (Info(5) of char 21 -
Info(5) of char 1) / 20 / 10.5 = the em advance. Measured 2026-09-15:
  ＭＳ 明朝 / ＭＳ ゴシック / ＭＳ Ｐ明朝 / 游明朝 / 游ゴシック, hint: all 1.000
  the same faces, no hint: σ 0.52-0.55 π 0.51-0.64 α 0.53-0.56 Я 0.68-0.75
                           × 0.56-0.61 ÷ 0.55 (= the ascii face, Century or TNR)
  メイリオ, hint: σ 0.632 π 0.621 α 0.607 Я 0.693 × 0.804 ÷ 0.804 (its own)
So: the hint routes the letter to the eastAsia face (S1419), and in the JIS
faces that glyph is full-width (S1416); Meiryo is proportional and stays on
the ascii routing (its own widths are not tabulated).
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/greek_cjk'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="567" w:bottom="1134" w:left="567" w:header="851" w:footer="992"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'
CH = 'σπαЯ×÷'


def rpr(ea, ascii_, hint):
    h = ' w:hint="eastAsia"' if hint else ''
    return f'<w:rPr><w:rFonts w:ascii="{ascii_}" w:eastAsia="{ea}" w:hAnsi="{ascii_}"{h}/><w:kern w:val="0"/><w:sz w:val="21"/></w:rPr>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for ea in ('ＭＳ 明朝', 'ＭＳ ゴシック', '游明朝', '游ゴシック', 'メイリオ', 'ＭＳ Ｐ明朝'):
        for ascii_ in ('Century', 'Times New Roman'):
            for hint in (True, False):
                paras = ''.join(f'<w:p><w:pPr><w:widowControl w:val="0"/><w:jc w:val="left"/>{rpr(ea, ascii_, hint)}</w:pPr><w:r>{rpr(ea, ascii_, hint)}<w:t>あ{ch * 20}あ</w:t></w:r></w:p>' for ch in CH)
                doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{paras}{SECT}</w:body></w:document>'
                at = OUT / f'{ea[:3]}_{ascii_[:3]}_{"hint" if hint else "nohint"}.docx'
                with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/document.xml', doc)
                d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
                try:
                    out = []
                    for pi, ch in enumerate(CH, 1):
                        r = d.Paragraphs(pi).Range; s = r.Start
                        x1 = d.Range(s + 1, s + 1).Information(5); x21 = d.Range(s + 21, s + 21).Information(5)
                        out.append(f'{ch}={(x21 - x1) / 20 / 10.5:.3f}')
                    print(f'{ea[:4]:5s} {ascii_[:3]} {"hint  " if hint else "nohint"} ' + ' '.join(out), flush=True)
                finally:
                    d.Close(False)
finally:
    app.Quit()
