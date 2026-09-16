# -*- coding: utf-8 -*-
"""Where does Word draw, and how many characters does it fit, when a paragraph has a NEGATIVE
left indent (leftChars<0) on a `docGrid type="linesAndChars"` page?

technical__9e4d04b4 (compat 11, linePitch 291, charSpace -3531 -> pitch 9.638, text width 481.85
= exactly 50 cells): its first paragraph carries ind leftChars="-100" left="-193" and holds 51
characters on ONE line — Word draws the first character AT THE LEFT MARGIN (x=71.25, not the
indented 61.25) and lets the 51st character end 8.25pt past the right margin. Oxi draws from the
indented origin and wraps the 51st, so every following block is one line (14pt) low.

Sheet: one paragraph of N full-width kana, ind leftChars swept. Readout: x of the first and last
character of line 1, how many characters line 1 holds, and the paragraph's line count.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/negindent'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
KANA = 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをんアイウエオカキクケコサシスセソタチツテト'


def settings(compat):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')


def document(n, left_chars, char_space, grid_type):
    left_tw = int(round(left_chars / 100.0 * 193))
    ind = '' if left_chars == 0 else f'<w:ind w:leftChars="{left_chars}" w:left="{left_tw}"/>'
    rpr = '<w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:hint="eastAsia"/></w:rPr>'
    body = (f'<w:p><w:pPr>{ind}{rpr}</w:pPr><w:r>{rpr}<w:t>{KANA[:n]}</w:t></w:r></w:p>'
            f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t>あと</w:t></w:r></w:p>')
    grid = '' if grid_type == 'none' else f'<w:docGrid w:type="{grid_type}" w:linePitch="291" w:charSpace="{char_space}"/>'
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1418" w:right="851" w:bottom="1134" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>'
            f'<w:cols w:space="425"/>{grid}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    if os.environ.get('ARMS3'):
        arms = [(51, 0, -3531, 'linesAndChars', c) for c in (11, 14, 15)]
        arms += [(52, -150, -3531, 'linesAndChars', c) for c in (11, 15)]
        arms += [(51, 0, -2048, 'linesAndChars', 11), (51, 0, -2048, 'linesAndChars', 15), (51, 0, 1024, 'linesAndChars', 11), (51, 0, 1024, 'linesAndChars', 15)]
    elif os.environ.get('ARMS2'):
        arms = [(n, lc, -3531, 'linesAndChars', c) for c in (11, 12, 14, 15) for lc in (-100, -200, -300) for n in (52,)]
        arms += [(52, 100, -3531, 'linesAndChars', 11), (52, 200, -3531, 'linesAndChars', 11), (52, 50, -3531, 'linesAndChars', 11), (52, -50, -3531, 'linesAndChars', 11)]
    else:
        arms = []
        for lc in (0, -100, -200, 100):
            for n in (50, 51, 52):
                arms.append((n, lc, -3531, 'linesAndChars', 11))
        arms += [(51, -100, -3531, 'linesAndChars', 15), (51, -100, 0, 'linesAndChars', 11), (51, -100, -3531, 'lines', 11), (51, -100, -3531, 'none', 11), (52, -200, -3531, 'linesAndChars', 15)]
    for n, lc, cs, gt, compat in arms:
        at = OUT / f'n{n}_lc{lc}_cs{cs}_{gt}_c{compat}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', settings(compat)); z.writestr('word/document.xml', document(n, lc, cs, gt))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            p = d.Paragraphs(1).Range
            rows = [(round(d.Range(c, c).Information(5), 2), round(d.Range(c, c).Information(6), 2)) for c in range(p.Start, p.End - 1)]
            y0 = rows[0][1]
            line1 = [r for r in rows if r[1] == y0]
            nlines = len(set(r[1] for r in rows))
            print(f'n={n} leftChars={lc:5} charSpace={cs:6} {gt:13} compat{compat} | first x={line1[0][0]} last x={line1[-1][0]} chars_line1={len(line1)} lines={nlines}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
