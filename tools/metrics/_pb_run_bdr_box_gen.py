# -*- coding: utf-8 -*-
"""Where does Word paint a run's w:bdr / w:shd box when the line is taller
than the run (docGrid pitch, exact, multiple)?

ja/reports__5373dd8a: «非正規労働者激増から…» (BIZ UDPゴシック bold 16, w:bdr
sz 4 + w:shd FFFF00, docGrid lines) -- Word's yellow rect 90.98..111.86 around
the baseline 107.54 (-1.035 / +0.27 em, 1.305 em tall) on a 36pt grid line;
the 18pt «「労働力調査」…» heading the same (-1.033 / +0.273 em). Oxi drew the
box over the whole 36pt line.

Arms: one bordered + shaded run (and an unbordered shaded run) in a
paragraph of each line regime -- single, docGrid lines 360tw, exact 30pt,
multiple 2.0 -- for ＭＳ 明朝 / BIZ UDPゴシック / Arial at 10.5 and 16pt.
Readout from Word's PDF: the shading fill rect and the run's baseline.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/run_bdr_box'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
FONTS = {'mincho': 'ＭＳ 明朝', 'bizp': 'BIZ UDPゴシック', 'arial': 'Arial'}
REGIMES = {'single': '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
           'grid': '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
           'exact': '<w:spacing w:after="0" w:line="600" w:lineRule="exact"/>',
           'mult2': '<w:spacing w:after="0" w:line="480" w:lineRule="auto"/>'}
TEXT = {'mincho': '枠付き文字列テスト', 'bizp': '枠付き文字列テスト', 'arial': 'Bordered run text'}


def document(font, size, regime, bdr):
    f = FONTS[font]
    rpr = (f'<w:rFonts w:ascii="{f}" w:eastAsia="{f}" w:hAnsi="{f}"/><w:sz w:val="{int(size * 2)}"/>'
           + ('<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>' if bdr else '')
           + '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>')
    p = f'<w:p><w:pPr>{REGIMES[regime]}</w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t>{TEXT[font]}</w:t></w:r></w:p>'
    grid = '<w:docGrid w:type="lines" w:linePitch="720"/>' if regime == 'grid' else '<w:docGrid w:linePitch="360"/>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            f'<w:p><w:pPr>{REGIMES[regime]}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="{f}" w:eastAsia="{f}" w:hAnsi="{f}"/><w:sz w:val="{int(size * 2)}"/></w:rPr><w:t>前</w:t></w:r></w:p>'
            f'{p}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            f'w:header="720" w:footer="720" w:gutter="0"/>{grid}</w:sectPr></w:body></w:document>')


ARMS = [(f, s, r, b) for f in FONTS for s in (10.5, 16.0) for r in REGIMES for b in (True, False)]

if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for arm in ARMS:
            font, size, regime, bdr = arm
            name = f'{font}_{size:g}_{regime}_{"bdr" if bdr else "shd"}'
            at = os.path.abspath(os.path.join(OUT, f'{name}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/document.xml', document(*arm))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = fitz.open(pdf)
            rects = [[round(v, 2) for v in dr['rect']] for dr in P[0].get_drawings()
                     if dr.get('fill') and dr['fill'][2] < 0.3 and dr['fill'][0] > 0.7]
            spans = [(round(s['origin'][1], 2), s['font'], round(s['size'], 2)) for b in P[0].get_text('dict')['blocks']
                     for l in b.get('lines', []) for s in l['spans'] if s['text'].strip()]
            base = spans[1][0] if len(spans) > 1 else None
            r = rects[0] if rects else None
            rel = (round((base - r[1]) / size, 4), round((r[3] - base) / size, 4), round((r[3] - r[1]) / size, 4)) if r and base else None
            print(f'{name:28} base {base} prev {spans[0][0] if spans else None} shd {r} rel(top,bot,h)/em {rel} {spans[1][1] if len(spans) > 1 else ""}', flush=True)
    finally:
        app.Quit()
