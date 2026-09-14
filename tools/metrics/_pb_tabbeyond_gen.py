# -*- coding: utf-8 -*-
"""A right tab stop BEYOND the right margin: does the line wrap, or extend?

policies__07543a6b9776a1cf (jablindC50): its TOC styles carry Word's default
US-Letter TOC tab «right leader=dot pos=9360» (6.5in) on an A4 page whose text
width is 8505tw (425.25pt). Word sets every entry on ONE line (Info6 pitch
13.5-19.5pt down the page); Oxi wraps each entry's page number to a 2nd line,
doubling the TOC and pushing «3.2 投薬過誤の用語選択事例» to page 4.

Arms (A4, margins 1701/1701 -> text width 425.25pt, Century 10.5, one
paragraph «第1章　はじめに<tab>1» then A2):
  beyond_dot     right tab 9360 (468pt, +42.75 past the margin), leader dot
  beyond_plain   right tab 9360, no leader
  at_margin      right tab 8505 (= margin), leader dot
  far_beyond     right tab 12000 (600pt, past the PAGE edge 595pt), leader dot
  left_beyond    LEFT tab 9360, no leader
  long_text      beyond_dot with a long entry that itself nearly fills the line
Readout: Word line count of the entry (Info6 sweep per char), x of the page
number's first char (Information(5)); Oxi line count + x from the dump.
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/tabbeyond'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '</Types>')
SETTINGS15 = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
SETTINGS14 = SETTINGS15.replace('w:val="15"', 'w:val="14"')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
          '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="toc1"><w:name w:val="toc 1"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:uiPriority w:val="39"/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs><w:spacing w:before="120"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial Bold" w:hAnsi="Arial Bold"/><w:b/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="toc3"><w:name w:val="toc 3"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:uiPriority w:val="39"/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs><w:ind w:left="720"/></w:pPr></w:style>'
          '</w:styles>')
SECT = ('<w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/><w:docGrid w:type="linesAndChars" w:linePitch="387"/></w:sectPr>')


def entry(text, pos, kind='right', leader=True, ind=''):
    ld = ' w:leader="dot"' if leader else ''
    return (f'<w:p><w:pPr><w:tabs><w:tab w:val="{kind}"{ld} w:pos="{pos}"/></w:tabs>{ind}</w:pPr>'
            f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>12</w:t></w:r></w:p>')


plain = '<w:p><w:r><w:t>A</w:t></w:r></w:p>'
WH = '<w:rPr><w:noProof/><w:webHidden/></w:rPr>'
FIELD = ('<w:r>' + WH + '<w:fldChar w:fldCharType="begin"/></w:r><w:r>' + WH + '<w:instrText xml:space="preserve"> PAGEREF _Toc1 \h </w:instrText></w:r>'
         '<w:r>' + WH + '</w:r><w:r>' + WH + '<w:fldChar w:fldCharType="separate"/></w:r><w:r>' + WH + '<w:t>12</w:t></w:r><w:r>' + WH + '<w:fldChar w:fldCharType="end"/></w:r>')
TABS = '<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs><w:spacing w:before="120"/></w:pPr>'
def toc_like(hyper=True, field=True, webhidden=True, pstyle=None):
    tab = '<w:r>' + (WH if webhidden else '') + '<w:tab/></w:r>'
    num = FIELD if field else '<w:r>' + (WH if webhidden else '') + '<w:t>12</w:t></w:r>'
    if not webhidden:
        num = num.replace('<w:webHidden/>', '')
    inner = '<w:r><w:t>第2章　データの質</w:t></w:r>' + tab + num
    if hyper:
        inner = '<w:hyperlink w:anchor="_Toc1" w:history="1">' + inner + '</w:hyperlink>'
    ppr = f'<w:pPr><w:pStyle w:val="{pstyle}"/></w:pPr>' if pstyle else TABS
    return '<w:p>' + ppr + inner + '</w:p>'
LONG = '2.4.2 MedDRAコーディングにおける考慮事項と品質保証の確認および研修プログラムの構成要素について'
arms = {
    'beyond_dot': entry('第1章　はじめに', 9360),
    'beyond_plain': entry('第1章　はじめに', 9360, leader=False),
    'at_margin': entry('第1章　はじめに', 8505),
    'far_beyond': entry('第1章　はじめに', 12000),
    'left_beyond': entry('第1章　はじめに', 9360, kind='left', leader=False),
    'long_text': entry(LONG, 9360),
    'toc_full': toc_like(),
    'toc_nohyper': toc_like(hyper=False),
    'toc_nofield': toc_like(field=False),
    'toc_nowebhid': toc_like(webhidden=False),
    'style_toc1': '<w:p><w:pPr><w:pStyle w:val="toc1"/></w:pPr><w:r><w:t>第2章　データの質</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>12</w:t></w:r></w:p>',
    'st3_full': toc_like(pstyle='toc3'),
    'st3_field': toc_like(hyper=False, webhidden=False, pstyle='toc3'),
    'st3_hyper': toc_like(field=False, webhidden=False, pstyle='toc3'),
    'st1_full': toc_like(pstyle='toc1'),
    'c15_beyond_dot': (entry('第1章　はじめに', 9360), 'SETTINGS15'),
    'c15_st3_full': (toc_like(pstyle='toc3'), 'SETTINGS15'),
    'c15_at_margin': (entry('第1章　はじめに', 8505), 'SETTINGS15'),
    'c14_beyond_dot': (entry('第1章　はじめに', 9360), 'SETTINGS14'),
    'c15_ind720': (entry('第1章　はじめに', 9360, ind='<w:ind w:left="720"/>'), 'SETTINGS15'),
    'c15_ind1440': (entry('第1章　はじめに', 9360, ind='<w:ind w:left="1440"/>'), 'SETTINGS15'),
    'c15_rind360': (entry('第1章　はじめに', 9360, ind='<w:ind w:right="360"/>'), 'SETTINGS15'),
    'c15_left9360': (entry('第1章　はじめに', 9360, kind='left', leader=False), 'SETTINGS15'),
    'c15_pos8000': (entry('第1章　はじめに', 8000), 'SETTINGS15'),
    'style_toc3': '<w:p><w:pPr><w:pStyle w:val="toc3"/></w:pPr><w:r><w:t>2.4.1 データ収集</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>12</w:t></w:r></w:p>',
}


def document(x):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{plain}{x}{plain}{SECT}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try:
    app.Visible = False
except Exception:
    pass
try:
    for name, arm in arms.items():
        x, setk = arm if isinstance(arm, tuple) else (arm, None)
        at = OUT / f"{name}.docx"
        with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", DOC_RELS); z.writestr("word/styles.xml", STYLES)
            if setk: z.writestr("word/settings.xml", globals()[setk])
            z.writestr("word/document.xml", document(x))
        d = app.Documents.Open(str(at.resolve()), False, True)
        try:
            p = d.Paragraphs(2); s, e = p.Range.Start, p.Range.End
            ys = sorted({round(d.Range(k, k).Information(6), 2) for k in range(s, e - 1)})
            num = d.Range(e - 3, e - 3)  # first char of '12'
            nx = round(num.Information(5), 2)
            ys3 = round(d.Range(d.Paragraphs(3).Range.Start, d.Paragraphs(3).Range.Start).Information(6), 2)
        finally:
            d.Close(False)
        with tempfile.TemporaryDirectory() as t:
            dump = Path(t) / 'l.json'
            subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
            dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
        els = [el for el in dd['pages'][0]['elements'] if el.get('type') == 'text']
        oys = sorted({round(el['y'], 2) for el in els})
        mid = [el for el in els if oys and oys[0] < round(el['y'], 2) < oys[-1]]
        oy = sorted({round(el['y'], 2) for el in mid})
        ox = [round(el['x'], 2) for el in mid if el['text'].strip() == '12']
        print(f"=== {name:13s} WORD lines {len(ys)} y {ys} num_x {nx} next_y {ys3} | OXI lines {len(oy)} y {oy} num_x {ox} next_y {oys[-1] if oys else None}")
finally:
    app.Quit()
