# -*- coding: utf-8 -*-
"""Where does a `<w:br w:type="page"/>` take effect when it sits in the RUNS of the
paragraph whose `pPr` carries the section's `sectPr`?

reference__0ea3ec86 offset 15558: section 0's last paragraph is

    <w:pPr><w:sectPr …><w:pgNumType w:start="88"/>…</w:sectPr></w:pPr>
    <w:r><w:br w:type="page"/></w:r></w:p>

Word renders the cover on page 1, a BLANK page 2, and the next section's heading on
page 3 (its pgNumType restarts at 90, and 88 -> 90 is the blank 89). Oxi puts the cover
and the next section's body together on page 1 and emits its blank two pages later, which
is the -2 on 68 paragraphs and the whole markers-off failure of that document.

Sheet: a cover paragraph, then the section-terminating paragraph, then the next section's
body. Readout: the page count and, per page, the first line of text, so an empty page is
visible as a page with no text at all.

Arms: the break's position (in the sectPr paragraph's runs / in its own paragraph after /
absent) x the next section's type (continuous / nextPage) x the sectPr paragraph carrying
text of its own or not x pgNumType restart (present / absent).
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/sect_tail_br'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ 明朝" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
          '<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')

PGSZ = ('<w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1304" w:right="1021" w:bottom="1134" w:left="1021" w:header="680" w:footer="567" w:gutter="0"/>'
        '<w:cols w:space="440"/><w:docGrid w:type="linesAndChars" w:linePitch="411" w:charSpace="3194"/>')


def document(brpos, nexttype, sect_text, restart):
    n1 = '<w:pgNumType w:start="88"/>' if restart else ''
    n2 = '<w:pgNumType w:start="90"/>' if restart else ''
    sect0 = f'<w:sectPr>{PGSZ}{n1}</w:sectPr>'
    txt = '<w:r><w:t>節の最後の段落です。</w:t></w:r>' if sect_text else ''
    br = '<w:r><w:br w:type="page"/></w:r>'
    body = '<w:p><w:r><w:t>表紙です。</w:t></w:r></w:p>'
    if brpos == 'in':
        body += f'<w:p><w:pPr>{sect0}</w:pPr>{txt}{br}</w:p>'
    elif brpos == 'after':
        body += f'<w:p><w:pPr>{sect0}</w:pPr>{txt}</w:p><w:p>{br}</w:p>'
    else:
        body += f'<w:p><w:pPr>{sect0}</w:pPr>{txt}</w:p>'
    body += '<w:p><w:r><w:t>つぎの節の見出しです。</w:t></w:r></w:p>'
    body += '<w:p><w:r><w:t>つぎの節の本文です。</w:t></w:r></w:p>'
    sect1 = f'<w:sectPr><w:type w:val="{nexttype}"/>{PGSZ}{n2}</w:sectPr>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect1}</w:body></w:document>'


def pages_of(pdf):
    import pymupdf
    doc = pymupdf.open(pdf)
    out = []
    for p in doc:
        t = [ln for ln in p.get_text().split('\n') if ln.strip()]
        out.append(t[0][:20] if t else '(BLANK)')
    doc.close()
    return out


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = []
    for brpos in ('in', 'after', 'none'):
        arms.append((brpos, 'continuous', False, True))
    arms += [('in', 'nextPage', False, True), ('none', 'nextPage', False, True),
             ('after', 'nextPage', False, True), ('none', 'oddPage', False, True),
             ('in', 'oddPage', False, True), ('in', 'continuous', True, True),
             ('in', 'continuous', False, False), ('after', 'continuous', False, False)]
    for brpos, nexttype, sect_text, restart in arms:
        tag = f'br{brpos}_{nexttype}_t{int(sect_text)}_r{int(restart)}'
        at = OUT / f'{tag}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', STYLES)
            z.writestr('word/document.xml', document(brpos, nexttype, sect_text, restart))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        pdf = str((OUT / f'{tag}.pdf').resolve())
        try:
            d.ExportAsFixedFormat(pdf, 17)
        finally:
            d.Close(False)
        pp = pages_of(pdf)
        print(f'br={brpos:5} next={nexttype:10} sectText={int(sect_text)} restart={int(restart)} | pages={len(pp)} {pp}', flush=True)
finally:
    app.Quit()
