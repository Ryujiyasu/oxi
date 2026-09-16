# -*- coding: utf-8 -*-
"""Does a table style's top-level <w:rPr><w:sz/> reach the cell text?

policies__07543a6b9776a1cf: Table Grid carries <w:rPr><w:sz w:val="20"/></w:rPr>,
the document's Normal (after linkStyles: the template's, no sz) leaves sz to
docDefaults 21 -- yet Word measures every cell character at 10.5, Oxi sets them
at 10. Arms: table style sz 20 with Normal {no sz | sz 21 | sz 24}, a cell
paragraph in a root style (no basedOn, no sz), and the same sheet under
linkStyles. Readout: Font.Size of the cell text and of a body paragraph.
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/tblstyle_sz'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
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


import os
OVERRIDE = os.environ.get('OVERRIDE', '0') == '1'   # overrideTableStyleFontSizeAndJustification


def settings(link):
    ov = '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>' if OVERRIDE else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + ('<w:linkStyles/>' if link else '') + '<w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>' + ov + '</w:compat></w:settings>')


def styles(normal_sz):
    nsz = f'<w:rPr><w:sz w:val="{normal_sz}"/></w:rPr>' if normal_sz else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="ＭＳ 明朝" w:hAnsi="Arial"/><w:kern w:val="2"/><w:sz w:val="21"/><w:lang w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            f'<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/></w:pPr>{nsz}</w:style>'
            '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Default"><w:name w:val="Default"/><w:pPr><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/></w:pPr><w:rPr><w:color w:val="000000"/></w:rPr></w:style>'
            '<w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
            '<w:style w:type="table" w:styleId="ac"><w:name w:val="Table Grid"/><w:basedOn w:val="a1"/><w:pPr><w:spacing w:after="120"/><w:jc w:val="both"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
            '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>'
            '</w:styles>')


SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992"/><w:docGrid w:linePitch="360"/></w:sectPr>'
DOC = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
       '<w:p><w:r><w:t>本文の段落です。</w:t></w:r></w:p>'
       '<w:tbl><w:tblPr><w:tblStyle w:val="ac"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
       '<w:tblGrid><w:gridCol w:w="4819"/><w:gridCol w:w="4819"/></w:tblGrid>'
       '<w:tr><w:tc><w:tcPr><w:tcW w:w="4819" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>標準のセル</w:t></w:r></w:p></w:tc>'
       '<w:tc><w:tcPr><w:tcW w:w="4819" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="Default"/></w:pPr><w:r><w:t>Default のセル</w:t></w:r></w:p></w:tc></w:tr>'
       '<w:tr><w:tc><w:tcPr><w:tcW w:w="4819" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>直書き jc のセル</w:t></w:r></w:p></w:tc>'
       '<w:tc><w:tcPr><w:tcW w:w="4819" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/><w:szCs w:val="21"/></w:rPr><w:t>太字のセル</w:t></w:r></w:p></w:tc></w:tr>'
       '</w:tbl><w:p><w:r><w:t>表のあとの段落。</w:t></w:r></w:p>'
       f'{SECT}</w:body></w:document>')

import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for link in (False, True):
        for nsz in (None, 21, 24):
            name = f'{"link" if link else "nolink"}_normal{nsz or "none"}{"_ov" if OVERRIDE else ""}'
            at = OUT / f'{name}.docx'
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
                z.writestr('word/styles.xml', styles(nsz)); z.writestr('word/settings.xml', settings(link)); z.writestr('word/document.xml', DOC)
            d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
            try:
                body = d.Paragraphs(1).Range
                t = d.Tables(1)
                cells = [(r, c, t.Cell(r, c).Range) for r in (1, 2) for c in (1, 2)]
                out = f'{name:20} body={body.Font.Size} '
                for r, c, rg in cells:
                    ch = d.Range(rg.Start, rg.Start + 1)
                    out += f'| c{r}{c} {ch.Font.Size} jc={rg.Paragraphs(1).Alignment} sa={rg.Paragraphs(1).SpaceAfter} '
                print(out, flush=True)
            finally:
                d.Close(False)
finally:
    app.Quit()
