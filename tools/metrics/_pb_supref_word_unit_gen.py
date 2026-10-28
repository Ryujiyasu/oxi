# -*- coding: utf-8 -*-
"""Is «Potton» + superscript «18» + «.» one unbreakable word at the line end?

en/legal__00081e80 p2 last line: «… Job v. Potton18.» -- Word (this PC) moves
the whole «Potton18.» to the next line; Oxi kept «Potton18» and wrapped «.»
alone (S899 flushes the word at the vertAlign seam, which also makes the seam
a break opportunity). The footnote reference then lands on the next page in
Word, not in Oxi.

Arms: filler length n = 0..7 extra «x» characters ahead of the tail so the tail
crosses the margin at different points; tail kinds: footnote reference / plain
superscript run (vertAlign) / no superscript (control). Readout per arm: line
(Information(10), first-character line number) of «Potton», of the
superscript, and of the final «.».
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/suprefunit'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>')
FN = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes {W}>'
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
      '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
      '<w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> note</w:t></w:r></w:p></w:footnote>'
      '</w:footnotes>')
R = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr>'
RS = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:vertAlign w:val="superscript"/><w:sz w:val="24"/></w:rPr>'
BASE = 'sec. 27; Denys v. Shuckburgh; Bentley v. Bates; Jacobs v. Seward; Job v. '


def document(n, kind):
    tail = f'<w:r>{R}<w:t>Potton</w:t></w:r>'
    if kind == 'fn':
        tail += f'<w:r>{RS}<w:footnoteReference w:id="1"/></w:r>'
    elif kind == 'sup':
        tail += f'<w:r>{RS}<w:t>18</w:t></w:r>'
    tail += f'<w:r>{R}<w:t>.</w:t></w:r>'
    body = (f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r>{R}<w:t xml:space="preserve">{"x" * n}{BASE}</w:t></w:r>{tail}</w:p>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1008" w:bottom="1440" w:left="1872" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for kind in ('fn', 'sup', 'none'):
        for n in range(0, 8):
            at = os.path.join(OUT, f'{kind}_n{n}.docx')
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS)
                z.writestr('word/footnotes.xml', FN)
                z.writestr('word/document.xml', document(n, kind))
            d = app.Documents.Open(os.path.abspath(at), ReadOnly=True)
            try:
                text = d.Paragraphs(1).Range.Text
                start = d.Paragraphs(1).Range.Start
                i = text.index('Potton')
                def line(k):
                    return d.Range(start + k, start + k).Information(10)
                end = len(text.rstrip('\r')) - 1
                print(f'{os.path.basename(at):14} Potton L{line(i)}  sup L{line(i + 6)}  dot L{line(end)}  x_dot={d.Range(start + end, start + end).Information(5):.1f}', flush=True)
            finally:
                d.Close(False)
finally:
    app.Quit()
