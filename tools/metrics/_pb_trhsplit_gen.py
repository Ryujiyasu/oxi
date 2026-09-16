# -*- coding: utf-8 -*-
"""Does Word split a multi-cell row whose atLeast trHeight is BINDING (content shorter
than the minimum) when the minimum does not fit above the page bottom?

policies__07543a6b9776a1cf p9: row 1 (trHeight 1439 atLeast = 72pt, three cells, the
tallest ~3 lines, vAlign center) starts at 714.75 with ~55pt of room; Word puts the
first lines on p9 and continues on p10. Oxi's S754 discriminator ("Word pushes any
trHeight row whole") moved it whole (+1 x3).

Sheet: A4, margins 1134 (content bottom 785.2), MS Mincho 10.5, 18pt grid, 34 filler
lines (bottom 668.7), a spacer with EXACT height X, then a 2-column table whose row 1
carries trHeight TRH atLeast and N one-line paragraphs per cell (N=3 -> 48pt content,
binding; N=5 -> 90pt, non-binding), vAlign center/top. Readout per arm: page / y of
row 1 line 1, row 1 last line, row 2.

  TRH=1440 N=3 VALIGN=center python _pb_trhsplit_gen.py
"""
import os, zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/trhsplit'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'
RPR = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:kern w:val="2"/><w:sz w:val="21"/></w:rPr>'
TRH = int(os.environ.get('TRH', '1440'))
N = int(os.environ.get('N', '3'))
VALIGN = os.environ.get('VALIGN', 'center')
FLOAT = os.environ.get('FLOAT', '0') == '1'   # floating table (tblpPr)
SB = int(os.environ.get('SB', '0'))          # spacing before (twips) on paragraph SBPARA of cell 1
SBPARA = int(os.environ.get('SBPARA', '0'))
XS = [int(v) for v in os.environ.get('XS', ','.join(str(i) for i in range(26, 91, 4))).split(',')]


def para(text, exact=None):
    sp = f'<w:spacing w:line="{int(exact * 20)}" w:lineRule="exact"/>' if exact else ''
    return f'<w:p><w:pPr>{sp}<w:widowControl w:val="0"/>{RPR}</w:pPr><w:r>{RPR}<w:t>{text}</w:t></w:r></w:p>'


def cell(i, n):
    def ppr(k):
        return f'<w:spacing w:before="{SB}"/>' if (SB and i == 1 and k + 1 == SBPARA) else ''
    ps = ''.join(f'<w:p><w:pPr>{ppr(k)}{RPR}</w:pPr><w:r>{RPR}<w:t>セル{i}の{k + 1}行目</w:t></w:r></w:p>' for k in range(n))
    return f'<w:tc><w:tcPr><w:tcW w:w="4819" w:type="dxa"/><w:vAlign w:val="{VALIGN}"/></w:tcPr>{ps}</w:tc>'


def table():
    r1 = f'<w:tr><w:trPr><w:trHeight w:val="{TRH}"/></w:trPr>{cell(1, N)}{cell(2, 1)}</w:tr>'
    r2 = f'<w:tr>{cell(3, 1)}{cell(4, 1)}</w:tr>'
    return ('<w:tbl><w:tblPr>' + ('<w:tblpPr w:leftFromText="142" w:rightFromText="142" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="center" w:tblpY="1"/>' if FLOAT else '') + '<w:tblW w:w="9638" w:type="dxa"/><w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            f'</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4819"/><w:gridCol w:w="4819"/></w:tblGrid>{r1}{r2}</w:tbl>')


def document(x):
    body = ''.join(para(f'埋め草の行 {i + 1}') for i in range(34)) + (para('隙間', exact=x) if x > 0 else '') + table() + para('表のあとの段落')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{SECT}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for x in XS:
        at = OUT / f'trh{TRH}_n{N}_{VALIGN}{"_float" if FLOAT else ""}{"_sb%d_%d" % (SB, SBPARA) if SB else ""}_{x:02d}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/document.xml', document(x))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            t = d.Tables(1)
            c1 = t.Cell(1, 1).Range; l1 = d.Range(c1.Start, c1.Start)
            pn = t.Cell(1, 1).Range.Paragraphs(N).Range; ln = d.Range(pn.Start, pn.Start)
            c2 = t.Cell(2, 1).Range; r2 = d.Range(c2.Start, c2.Start)
            room = 785.2 - 668.7 - x
            sbp = ''
            if SB:
                q = t.Cell(1, 1).Range.Paragraphs(SBPARA).Range; qq = d.Range(q.Start, q.Start); sbp = f' para{SBPARA} p{int(qq.Information(3))} y={qq.Information(6):.2f} |'
            print(f'trh{TRH} n{N} {VALIGN}{" float" if FLOAT else ""} X={x:2d} room={room:5.1f}{sbp} row1 line1 p{int(l1.Information(3))} y={l1.Information(6):.2f} line{N} p{int(ln.Information(3))} y={ln.Information(6):.2f} | row2 p{int(r2.Information(3))} y={r2.Information(6):.2f}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
