# -*- coding: utf-8 -*-
"""Does a middle dot lend half an em to a justified compressPunctuation cell line?

ja/educational__244a3664: «・　夢としてはじめから就きたいものではなかったが…» in a
cell whose first line holds 19.66 em; the 20th character overflows by 3.55pt.
Word breaks after 19 (the dot and the fullwidth space keep their full advance);
Oxi's S1209 pool credits the dot half an em and keeps 20.

Arms (compat 15, jc both, compressPunctuation, 10.5pt ＭＳ 明朝, inner width 19.66
em): where the dot sits, and a 、 / 「 control. Readout: chars on line 1.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/cell_middot_lend'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
REST = 'としてはじめから就きたいものではなかったがやりがいは後からきた'
ARMS = {
    'dot_sp_start': '・　夢' + REST,
    'dot_start': '・夢と' + REST,
    'sp_start': '一　夢' + REST,
    'dot_mid': '夢と・' + REST,
    'dot_sp_mid': '夢・　' + REST,
    'comma_mid': '夢と、' + REST,
    'kanji': '一二夢' + REST,
}


def settings(c, bal=False):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat>' + ('<w:balanceSingleByteDoubleByteWidth/>' if bal else '') +
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{c}"/></w:compat></w:settings>')


def document(text):
    rpr = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'
    p = f'<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>'
    tc = '<w:tc><w:tcPr><w:tcW w:w="4327" w:type="dxa"/></w:tcPr>' + p + '</w:tc>'
    b = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>' for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    tbl = (f'<w:tbl><w:tblPr><w:tblW w:w="4327" w:type="dxa"/><w:tblBorders>{b}</w:tblBorders>'
           '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="4327"/></w:tblGrid><w:tr>{tc}</w:tr></w:tbl>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body><w:p/>{tbl}<w:p/>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
            'w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for c, bal in ((15, False), (14, False), (15, True), (14, True)):
            for name, text in ARMS.items():
                at = os.path.abspath(os.path.join(OUT, f'c{c}{"b" if bal else ""}_{name}.docx'))
                with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                    z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', settings(c, bal))
                    z.writestr('word/document.xml', document(text))
                pdf = at[:-5] + '.pdf'
                d = app.Documents.Open(at, ReadOnly=True)
                try:
                    d.ExportAsFixedFormat(pdf, 17)
                finally:
                    d.Close(False)
                P = pymupdf.open(pdf)[0]
                lines = {}
                for b in P.get_text('rawdict')['blocks']:
                    for l in b.get('lines', []):
                        for s in l['spans']:
                            for ch in s['chars']:
                                lines.setdefault(round(s['origin'][1], 1), []).append(ch['c'])
                ys = [y for y in sorted(lines) if ''.join(lines[y]).strip()]
                l1 = ''.join(lines[ys[0]]).rstrip() if ys else ''
                print(f'c{c}{"b" if bal else " "} {name:12} line1 n {len(l1):2} {l1}', flush=True)
    finally:
        app.Quit()
