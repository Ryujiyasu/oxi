# -*- coding: utf-8 -*-
"""When does a default tab whose next stop lies past the line end clamp, and when
does it wrap to the next line?

S1044 measured a clamp with 28.8pt of room left (forms__00042714) and a wrap when
the line is already at the boundary; `_pb_tab_float_cross_gen.py` shows Word
wrapping the 14th tab with 16.3pt left (12 TNR-10 tabs, then a 13th to 525, then
the 14th). Arms: 12 tabs (the 13th starts at 489) + one tab + a tail, with the
right indent swept so the room between 489 and the line end runs 6..34pt; tail =
«X» or «XXXXXXXX». compat 15. Readout: Word COM Information(5)/(6) of the 13th tab
and of the tail.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/tab_room'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
def SET(c):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="720"/><w:compat>'
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{c}"/></w:compat></w:settings>')
RPR = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr>'


def document(right_pt, tail, prefix='', before=''):
    tabs = ''.join(f'<w:r>{RPR}<w:tab/></w:r>' for _ in range(12)) + (f'<w:r>{RPR}<w:t>{before}</w:t></w:r>' if before else '') + f'<w:r>{RPR}<w:tab/></w:r>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            f'<w:p><w:pPr><w:ind w:right="{int(right_pt * 20)}"/></w:pPr><w:r>{RPR}<w:t>{prefix}</w:t></w:r>{tabs}<w:r>{RPR}<w:t>{tail}</w:t></w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1140" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        print('line end 541.3 - right; the 13th tab starts at 489')
        arms = [(15, '', 'X', room, b) for b in ('i', 'ii') for room in (6, 18, 30)]
        for c, pre, tail, room, b in arms:
                right = 541.3 - 489 - room
                name = f'c{c}_before_{b}_room{room}'
                at = os.path.abspath(os.path.join(OUT, name + '.docx'))
                with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                    z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', SET(c))
                    z.writestr('word/document.xml', document(right, tail, pre, b))
                d = app.Documents.Open(at, ReadOnly=True)
                try:
                    r = d.Paragraphs(1).Range
                    pos = [(d.Range(i, i + 1).Text.replace('\t', 'T'), round(d.Range(i, i).Information(5), 1), round(d.Range(i, i).Information(6), 1))
                           for i in range(r.Start, r.End - 1)]
                    t13 = [p for p in pos if p[0] == 'T'][12]
                    tl = [p for p in pos if p[0] == 'X'][0]
                    print(f'{name:18} right {right:5.1f}  tab13 {t13[1:]}  tail {tl[1:]}', flush=True)
                finally:
                    d.Close(False)
    finally:
        app.Quit()
