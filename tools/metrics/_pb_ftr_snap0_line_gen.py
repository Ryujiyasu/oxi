# -*- coding: utf-8 -*-
"""Where does Word put a one-line footer's baseline when the footer paragraph
says snapToGrid=0 under a linesAndChars grid?

ja/creative__13152ea1: footer «５月２１日（日）は…» (HG創英角ﾎﾟｯﾌﾟ体 14, w:bdr sz4,
footer style snapToGrid 0, docGrid linesAndChars 375 / charSpace 194, footer
992tw): Word baseline 787.90, Oxi 792.30 (+4.4).

Arms: the same footer paragraph with / without the run border, snapToGrid 0 /
1, grid linesAndChars / lines / none, font HG創英角ﾎﾟｯﾌﾟ体 / ＭＳ 明朝 at 14pt.
Readout: the footer line's baseline from Word's PDF.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/ftr_snap0_line'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
         '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
FONTS = {'pop': 'HG創英角ﾎﾟｯﾌﾟ体', 'mincho': 'ＭＳ 明朝'}
GRIDS = {'lac': '<w:docGrid w:type="linesAndChars" w:linePitch="375" w:charSpace="194"/>',
         'lines': '<w:docGrid w:type="lines" w:linePitch="375"/>', 'none': '<w:docGrid w:linePitch="360"/>'}


def footer(font, bdr, snap):
    f = FONTS[font]
    rpr = (f'<w:rFonts w:ascii="{f}" w:eastAsia="{f}" w:hAnsi="{f}" w:hint="eastAsia"/><w:sz w:val="28"/>'
           + ('<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>' if bdr else ''))
    ppr = '<w:jc w:val="center"/>' + ('' if snap else '<w:snapToGrid w:val="0"/>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {W}><w:p><w:pPr>{ppr}</w:pPr>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t>５月２１日（日）は運動会です。</w:t></w:r></w:p></w:ftr>')


def document(grid):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body><w:p><w:r><w:t>本文</w:t></w:r></w:p>'
            '<w:sectPr><w:footerReference w:type="default" r:id="rId7"/><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>'
            f'{GRIDS[grid]}</w:sectPr></w:body></w:document>')


ARMS = [(f, b, s, g) for f in FONTS for b in (True, False) for s in (False, True) for g in GRIDS]

if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for f, b, s, g in ARMS:
            name = f'{f}_{"bdr" if b else "plain"}_{"snap1" if s else "snap0"}_{g}'
            at = os.path.abspath(os.path.join(OUT, f'{name}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', SETTINGS)
                z.writestr('word/footer1.xml', footer(f, b, s)); z.writestr('word/document.xml', document(g))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = fitz.open(pdf)
            ys = [round(sp['origin'][1], 2) for bl in P[0].get_text('dict')['blocks'] for l in bl.get('lines', [])
                  for sp in l['spans'] if '運動会' in sp['text'] or sp['text'].startswith('５')]
            print(f'{name:28} baseline {ys[:1]}', flush=True)
    finally:
        app.Quit()
