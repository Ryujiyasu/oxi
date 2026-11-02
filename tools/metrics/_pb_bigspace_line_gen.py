# -*- coding: utf-8 -*-
"""Does a whitespace-only run taller than the text grow its line?

ja/administrative__0ba21b0c (table cell, centred, line atLeast 0): «　» at 24pt
bold + « » at 14pt + «７・８月の主な予定» at 9pt. Word's COM tops: this line
181.5, the next paragraph 197.25 (15.75 for the line), as if the 24pt space were
not there; Oxi grew the line to the space and pushed the cell 18pt down.

Arms: space char (U+3000 / U+0020 / letter control «Ａ») x position (lead / mid /
trail) x jc (left / center) x context (body / cell). Each document: paragraph P
(the arm) then a 9pt paragraph «次»; readout = COM Information(6) difference.
"""
import os, sys, zipfile, itertools
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/bigspace_line'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
FONT = 'ＭＳ 明朝'


def run(text, sz):
    return (f'<w:r><w:rPr><w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/><w:sz w:val="{sz}"/></w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t></w:r>')


def para(runs, jc):
    return f'<w:p><w:pPr><w:spacing w:after="0" w:line="0" w:lineRule="atLeast"/><w:jc w:val="{jc}"/></w:pPr>{runs}</w:p>'


def arm_runs(ch, where):
    big = run(ch, 48)
    small = run('七八月の予定', 18)
    return {'lead': big + small, 'mid': run('七八月', 18) + big + run('の予定', 18), 'trail': small + big}[where]


def document(ch, where, jc, ctx):
    p = para(arm_runs(ch, where), jc) + para(run('次の段落', 18), 'left')
    if ctx == 'cell':
        p = ('<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>'
             f'<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>{p}</w:tc></w:tr></w:tbl><w:p/>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{p}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for ch, where, jc, ctx in itertools.product(('　', ' ', 'Ａ'), ('lead', 'mid', 'trail'), ('left', 'center'), ('body', 'cell')):
            tag = {'　': 'ideo', ' ': 'sp', 'Ａ': 'A'}[ch]
            name = f'{tag}_{where}_{jc}_{ctx}'
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/document.xml', document(ch, where, jc, ctx))
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                y = [d.Range(d.Paragraphs(i).Range.Start, d.Paragraphs(i).Range.Start).Information(6) for i in (1, 2)]
            finally:
                d.Close(False)
            print(f'{name:24} line {y[1] - y[0]:6.2f}', flush=True)
    finally:
        app.Quit()
