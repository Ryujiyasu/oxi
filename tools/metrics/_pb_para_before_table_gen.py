# -*- coding: utf-8 -*-
"""How tall is the paragraph IMMEDIATELY BEFORE a table on a no-type docGrid?

technical__5175ec20 p8 (no-type docGrid linePitch=360 = 18pt, body ＭＳ 明朝 10.5, tables
tblBorders sz=4 with no tblCellMar): comparing Word's Information(6) against Oxi's dump
line-box tops paragraph by paragraph, everything INSIDE a table agrees within 0.55 and the
gap leaving a table agrees within 1.2, but the paragraph that sits directly above a table
is short in Oxi twice over:

    para 425 '変更前'  Word 475.50 -> table rule 491.11 = 15.61 ; Oxi 475.87 -> 489.87 = 14.00
    para 436 '変更後'  Word 581.25 -> table rule 597.10 = 15.85 ; Oxi 580.72 -> 595.22 = 14.50

An ORDINARY paragraph in the same section is 18.75/19.50 (the 18pt grid pitch plus Info6's
0.75 quantisation), so Word shortens the last paragraph before a table too — just not as
far as Oxi does. The two occurrences are the ~3pt that lets Oxi fit the 3.18 heading on
page 8 where Word pushes it to page 9, and the same shape is all three remaining
markers-off JA failures.

Sheet: three body paragraphs, then a one-row table, then a closing paragraph. Readout:
Information(6) of every paragraph (so the ordinary paragraph height and the last one are
measured the same way) AND the PDF's top rule for the table (so the table's own top is
measured font-independently).

Arms: docGrid (no-type 360 / no-type 240 / type=lines 360 / absent) x table border weight
(sz 4 / 12) x tblInd (0 / 108) x body size (21 / 24).
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/para_before_table'); OUT.mkdir(parents=True, exist_ok=True)
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
def SETTINGS_FN(mode):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')


SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def styles(sz):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ 明朝" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style></w:styles>')


def table(bsz, ind):
    b = ''.join(f'<w:{t} w:val="single" w:sz="{bsz}" w:space="0" w:color="auto"/>'
                for t in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    return ('<w:tbl><w:tblPr><w:tblW w:w="8647" w:type="dxa"/>'
            f'<w:tblInd w:w="{ind}" w:type="dxa"/><w:tblBorders>{b}</w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2694"/><w:gridCol w:w="5953"/></w:tblGrid>'
            '<w:tr><w:trPr><w:trHeight w:val="451"/></w:trPr>'
            '<w:tc><w:tcPr><w:tcW w:w="2694" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>報告語</w:t></w:r></w:p></w:tc>'
            '<w:tc><w:tcPr><w:tcW w:w="5953" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>選択されたLLT</w:t></w:r></w:p></w:tc>'
            '</w:tr></w:tbl>')


def document(grid, bsz, ind, runsz=None):
    rpr = f'<w:rPr><w:sz w:val="{runsz}"/><w:szCs w:val="{runsz}"/></w:rPr>' if runsz else ''
    body = ''.join(f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t>本文の段落{i}です。</w:t></w:r></w:p>' for i in range(3))
    body += table(bsz, ind) + '<w:p><w:r><w:t>あとの段落です。</w:t></w:r></w:p>'
    if grid == 'none':
        g = ''
    elif grid == 'lines360':
        g = '<w:docGrid w:type="lines" w:linePitch="360"/>'
    else:
        g = f'<w:docGrid w:linePitch="{grid}"/>'
    sect = ('<w:sectPr><w:pgSz w:w="11907" w:h="16840" w:code="9"/><w:pgMar w:top="1440" w:right="1418" w:bottom="1440" w:left="1418" w:header="720" w:footer="720"/>'
            f'<w:cols w:space="425"/>{g}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


def top_rule(pdf):
    import pymupdf
    doc = pymupdf.open(pdf); p = doc[0]; ys = set()
    for dr in p.get_drawings():
        for it in dr['items']:
            if it[0] == 're' and it[1].height <= 2.0 and it[1].width > 50:
                ys.add(round(it[1].y0, 2))
            elif it[0] == 'l' and abs(it[1].y - it[2].y) < 0.5 and abs(it[1].x - it[2].x) > 50:
                ys.add(round(it[1].y, 2))
    doc.close()
    return min(ys) if ys else None


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = [('360', 21, 4, 108)]
    for grid in ('240', 'lines360', 'none'):
        arms.append((grid, 21, 4, 108))
    arms += [('360', 21, 12, 108), ('360', 21, 4, 0), ('360', 24, 4, 108),
             ('360', 24, 4, 108, 21), ('none', 24, 4, 108, 21), ('lines360', 24, 4, 108, 21),
             ('360', 24, 4, 108, 21, 14), ('360', 21, 4, 108, None, 14), ('none', 21, 4, 108, None, 14),
             ('240', 21, 4, 108, None, 14)]
    for arm in arms:
        grid, sz, bsz, ind = arm[:4]
        runsz = arm[4] if len(arm) > 4 else None
        mode = arm[5] if len(arm) > 5 else 15
        tag = f'g{grid}_sz{sz}_b{bsz}_i{ind}' + (f'_r{runsz}' if runsz else '') + f'_c{mode}'
        at = OUT / f'{tag}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS_FN(mode)); z.writestr('word/styles.xml', styles(sz))
            z.writestr('word/document.xml', document(grid, bsz, ind, runsz))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        pdf = str((OUT / f'{tag}.pdf').resolve())
        try:
            ys = []
            for i in range(1, 4):
                r = d.Paragraphs(i).Range
                ys.append(round(d.Range(r.Start, r.Start).Information(6), 2))
            d.ExportAsFixedFormat(pdf, 17)
        finally:
            d.Close(False)
        rule = top_rule(pdf)
        ordinary = round(ys[1] - ys[0], 2)
        last = round(rule - ys[2], 2) if rule else None
        print(f'grid={grid:9} c={mode} sz={sz} run={str(runsz):4} bsz={bsz:3} ind={ind:4} | p_y={ys} rule={rule} ordinary={ordinary} last_before_table={last}', flush=True)
finally:
    app.Quit()
