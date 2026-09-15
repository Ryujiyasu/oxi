# -*- coding: utf-8 -*-
"""MS PGothic / PMincho / Gothic kana advance by point size: does Word set them
at the face's hmtx fraction (size-independent em ratio) or at the embedded
bitmap strike's integer 96-dpi pixels?

policies__0af65d4597412f1b p2 (5): Word's line 1 holds 56 chars, Oxi 60 -- Oxi's
kana were ~0.83em (the MS UI Gothic fallback). Measured 2026-09-15 (Word COM,
Information(5) of char 0 and char 30 in a 30-char run, /30):
  ＭＳ Ｐゴシック  の = 1.000em, に 0.941, り 0.747, て 0.903 at every size 8..16pt
  ＭＳ Ｐ明朝      の = 0.950,   に 0.950, り 0.641, て 0.903
  ＭＳ ゴシック    all 1.000
i.e. the hmtx advances (GDI at 2048px gives 2048/1928/1528/1848 for PGothic),
not the bitmap strikes (at 12px GDI hints に to 11px = 0.917). 18pt+ rows wrap
(30 chars exceed the line) and are not readable this way.
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/pgothic_px'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="567" w:bottom="1134" w:left="567" w:header="851" w:footer="992"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'


def rpr(fam, sz):
    return f'<w:rPr><w:rFonts w:ascii="{fam}" w:eastAsia="{fam}" w:hAnsi="{fam}"/><w:kern w:val="0"/><w:sz w:val="{int(sz*2)}"/></w:rPr>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for fam in ('ＭＳ Ｐゴシック', 'ＭＳ Ｐ明朝', 'ＭＳ ゴシック'):
        for sz in (8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24):
            paras = ''.join(f'<w:p><w:pPr><w:widowControl w:val="0"/><w:jc w:val="left"/>{rpr(fam, sz)}</w:pPr><w:r>{rpr(fam, sz)}<w:t>{ch * 30}</w:t></w:r></w:p>' for ch in 'のにりて')
            doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{paras}{SECT}</w:body></w:document>'
            at = OUT / f'{fam[:3]}_{sz}.docx'
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/document.xml', doc)
            d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
            try:
                out = []
                for pi, ch in enumerate('のにりて', 1):
                    r = d.Paragraphs(pi).Range; s = r.Start
                    x0 = d.Range(s, s).Information(5); x30 = d.Range(s + 30, s + 30).Information(5)
                    y0 = d.Range(s, s).Information(6); y30 = d.Range(s + 29, s + 29).Information(6)
                    same = abs(y0 - y30) < 0.5
                    out.append(f'{ch}={(x30 - x0) / 30:.3f}' + ('' if same else '(wrap)'))
                print(f'{fam[:5]} {sz:5}pt ({sz * 96 / 72:5.2f}px) ' + ' '.join(out), flush=True)
            finally:
                d.Close(False)
finally:
    app.Quit()
