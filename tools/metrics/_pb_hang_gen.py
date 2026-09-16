# -*- coding: utf-8 -*-
"""Does a line-end 、 hang past the text edge (burasagari) or push its host down (oidashi)?

policies__07543a6b p28 (compat 15, doNotCompress, ＭＳ 明朝 10.5): a table cell 123.95pt
wide holds 11 full-width characters; Word breaks 「…服用した / 後、横紋筋…」 while Oxi
hangs the 、 past the edge and fits 「…服用した後、」. S506 recorded compat 12/14 hang,
15 oidashi, for a body paragraph.

Sheet: text = 11 chars + 「後、」 + more, in (a) a body paragraph whose width is set by
margins so that exactly 11 chars fit, (b) a table cell of the same inner width.
Arms: compat 14/15 × jc left/both × body/cell. Readout: chars on line 1.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/hang'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
RPR = '<w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="ＭＳ 明朝" w:hAnsi="Arial"/><w:kern w:val="2"/><w:sz w:val="21"/></w:rPr>'
TEXT = '患者は薬物と薬物相互作用がラベル表示された二、つの薬剤を服用した後横紋筋融解症を発現した。'  # 22 chars then 、 = the 23rd character, the 1st slot of line 3 at 11 per line
CSC = os.environ.get('CSC', 'doNotCompress')
WTW = int(os.environ.get('WTW', '2479'))   # inner width in twips (2479 = 11.8 chars of 10.5pt)


def settings(compat):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:characterSpacingControl w:val="{CSC}"/><w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')


def para(jc, text=TEXT):
    return f'<w:p><w:pPr><w:jc w:val="{jc}"/>{RPR}</w:pPr><w:r>{RPR}<w:t>{text}</w:t></w:r></w:p>'


def document(compat, jc, where):
    # inner width 11.8 chars: 123.95pt = 2479 twips
    if where == 'body':
        margins = (11906 - WTW) // 2
        body = para(jc)
        sect = f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="{margins}" w:bottom="1134" w:left="{11906 - WTW - margins}" w:header="851" w:footer="992"/><w:docGrid w:linePitch="360"/></w:sectPr>'
    else:
        cellw = WTW + 216
        body = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
                f'<w:tblGrid><w:gridCol w:w="{cellw}"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="{cellw}" w:type="dxa"/></w:tcPr>{para(jc)}</w:tc><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>{para(jc, "隣")}</w:tc></w:tr></w:tbl><w:p/>')
        sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992"/><w:docGrid w:linePitch="360"/></w:sectPr>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for compat in (14, 15):
        for jc in ('left', 'both'):
            for where in ('body', 'cell'):
                at = OUT / f'c{compat}_{jc}_{where}_{CSC}_w{WTW}.docx'
                with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
                    z.writestr('word/settings.xml', settings(compat)); z.writestr('word/document.xml', document(compat, jc, where))
                d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
                try:
                    pr = d.Paragraphs(1).Range if where == 'body' else d.Tables(1).Cell(1, 1).Range.Paragraphs(1).Range
                    lines = []; last = None; cur = ''
                    for k in range(pr.Start, pr.End - (0 if where == 'body' else 1)):
                        y = round(d.Range(k, k).Information(6), 2)
                        if y != last and cur:
                            lines.append(cur); cur = ''
                        last = y; cur += d.Range(k, k + 1).Text
                    if cur: lines.append(cur)
                    print(f'compat{compat} {jc:4} {where:4} {CSC:20} w{WTW} lines={lines}', flush=True)
                finally:
                    d.Close(False)
finally:
    app.Quit()
