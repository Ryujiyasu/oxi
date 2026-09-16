# -*- coding: utf-8 -*-
"""What does `w:linkStyles` (no attachedTemplate) replace from the local Normal.dotm?

policies__07543a6b9776a1cf carries `<w:linkStyles/>` and no attachedTemplate
relationship. Its own styles.xml says Normal = minorHAnsi + jc both, docDefaults
Times New Roman + minorEastAsia, theme minor Arial / ＭＳ Ｐゴシック -- yet Word's
truth is 游明朝, left-aligned: the values of THIS machine's Normal.dotm (Normal =
widowControl only, docDefaults = theme fonts, theme minor = 游明朝).

Sheet: a docx whose docDefaults (ＭＳ ゴシック 12, pPrDefault after 6), Normal
(ＭＳ 明朝 10.5, jc both, before 6) and a document-only style MyBody all differ
from Normal.dotm. Arms: no linkStyles / linkStyles. Readout through Word COM:
per-paragraph font / size / alignment / spacing and the resolved Normal style.
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/linkstyles'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')


def settings(link):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + ('<w:linkStyles/>' if link else '') + '<w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


STY = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
       '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:kern w:val="2"/><w:sz w:val="24"/><w:lang w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults>'
       '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/><w:spacing w:before="120"/></w:pPr><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:style>'
       '<w:style w:type="paragraph" w:styleId="mine"><w:name w:val="MyBody"/><w:basedOn w:val="a"/><w:pPr><w:ind w:left="420"/></w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:style></w:styles>')
SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'
DOC = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
       '<w:p><w:r><w:t>標準スタイルの段落です。両端揃えと段落前の間隔が文書側の指定。</w:t></w:r></w:p>'
       '<w:p><w:pPr><w:pStyle w:val="mine"/></w:pPr><w:r><w:t>文書独自スタイルの段落です。</w:t></w:r></w:p>'
       '<w:p><w:r><w:t>Latin text in the third paragraph.</w:t></w:r></w:p>'
       f'{SECT}</w:body></w:document>')
for link in (False, True):
    at = OUT / f'{"link" if link else "nolink"}.docx'
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
        z.writestr('word/styles.xml', STY); z.writestr('word/settings.xml', settings(link)); z.writestr('word/document.xml', DOC)

import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for name in ('nolink', 'link'):
        d = app.Documents.Open(str((OUT / f'{name}.docx').resolve()), ReadOnly=True)
        try:
            print('==', name, 'attached:', d.AttachedTemplate.Name)
            for i in range(1, 4):
                p = d.Paragraphs(i); r = p.Range
                print(f'  para{i} style={p.Style.NameLocal!r} font={r.Font.Name!r}/{r.Font.NameFarEast!r} {r.Font.Size} align={p.Alignment} sb={p.SpaceBefore} sa={p.SpaceAfter} left={p.LeftIndent} y={d.Range(r.Start, r.Start).Information(6):.2f}')
            st = d.Styles(-1)
            print('  Normal style: font', st.Font.Name, st.Font.NameFarEast, st.Font.Size, 'align', st.ParagraphFormat.Alignment, 'sb', st.ParagraphFormat.SpaceBefore, 'sa', st.ParagraphFormat.SpaceAfter)
        finally:
            d.Close(False)
finally:
    app.Quit()
