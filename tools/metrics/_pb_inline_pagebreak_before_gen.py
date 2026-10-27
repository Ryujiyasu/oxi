# -*- coding: utf-8 -*-
"""Does a paragraph that opens with an inline <w:br w:type="page"/> keep its
space-before at the top of the new page?

en/reports__00b179e8 p10: Heading1 (before 360) = «<w:br w:type="page"/>II.
Undergraduate Education», compat 15. Word sets the heading at the body top
(71.7); Oxi kept the 18pt (89.3) because page_break_before_from_inline is
excluded from the compat-15 suppression.

Arms: inline break vs pageBreakBefore, compat 15 / 14, Latin (Arial) / CJK
(ＭＳ 明朝), with and without a 2-line header. Readout: y of the paragraph's
first visible character (Information(6)) on its page.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/inlinepb'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
         '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>')


def settings(compat):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')


def rpr(face):
    return f'<w:rPr><w:rFonts w:ascii="{face}" w:eastAsia="{face}" w:hAnsi="{face}"/><w:sz w:val="18"/></w:rPr>'


def header(face, lines):
    body = ''.join(f'<w:p><w:r>{rpr(face)}<w:t>Header line {i}</w:t></w:r></w:p>' for i in range(lines))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{body or "<w:p/>"}</w:hdr>'


def document(kind, face, text, hdr):
    r = rpr(face)
    if kind == 'inline':
        heading = (f'<w:p><w:pPr><w:spacing w:before="360" w:after="180"/>{r}</w:pPr>'
                   f'<w:r>{r}<w:br w:type="page"/></w:r><w:r>{r}<w:t>{text}</w:t></w:r></w:p>')
    else:
        heading = (f'<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="360" w:after="180"/>{r}</w:pPr>'
                   f'<w:r>{r}<w:t>{text}</w:t></w:r></w:p>')
    href = '<w:headerReference w:type="default" r:id="rId8"/>' if hdr else ''
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            f'<w:p><w:pPr>{r}</w:pPr><w:r>{r}<w:t>{text} first page</w:t></w:r></w:p>'
            f'{heading}'
            f'<w:p><w:pPr>{r}</w:pPr><w:r>{r}<w:t>{text} after</w:t></w:r></w:p>'
            f'<w:sectPr>{href}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="721" w:right="721" w:bottom="721" w:left="721" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>')


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for compat in (15, 14):
        for face, text in (('Arial', 'Heading'), ('ＭＳ 明朝', '見出し')):
            for hdr in (0, 2):
                for kind in ('inline', 'pbb'):
                    at = os.path.join(OUT, f'c{compat}_{"lat" if face == "Arial" else "cjk"}_h{hdr}_{kind}.docx')
                    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                        z.writestr('word/_rels/document.xml.rels', DRELS)
                        z.writestr('word/settings.xml', settings(compat))
                        z.writestr('word/header1.xml', header(face, hdr))
                        z.writestr('word/document.xml', document(kind, face, text, hdr))
                    d = app.Documents.Open(os.path.abspath(at), ReadOnly=True)
                    try:
                        p = d.Paragraphs(2).Range
                        t = p.Text
                        k = next(i for i, c in enumerate(t) if c not in '\x0c\r')
                        c = d.Range(p.Start + k, p.Start + k)
                        page = c.Information(3); y = c.Information(6)
                        nxt = d.Paragraphs(3).Range
                        ny = d.Range(nxt.Start, nxt.Start).Information(6)
                        print(f'c{compat} {face:6} hdr={hdr} {kind:6} page={page} y={y:.2f} next_y={ny:.2f}', flush=True)
                    finally:
                        d.Close(False)
finally:
    app.Quit()
