# -*- coding: utf-8 -*-
"""An EMPTY paragraph with space-before on a linesAndChars grid: how tall?

policies__07543a6b9776a1cf (jablindC50, docGrid linesAndChars 387 = 19.35pt,
docDefaults TNR 12 / minorEastAsia, lang eastAsia en-US): the cover page stacks
«公表版1.0» (Century 18.5, before 9) + three EMPTY paragraphs whose font is the
unknown «ＭＳ明朝,Bold» (fontTable altName ＭＳ 明朝, notTrueType) at 18.5pt
with before 9, then «2018年 6月». Word: each empty = 2 grid rows (39) + its
9pt before = 48 per paragraph; Oxi: 38.75 per paragraph -- 9.25 short each,
28.5 over the page, two paragraphs pulled up from page 2.

Arms (same section; paragraphs A / E1 / E2 / E3 / B; readout = Information(6)
of every paragraph, so E-heights and befores separate):
  real       as in the doc: E font «ＭＳ明朝,Bold» (+ fontTable entry), before 9, bold
  lit_mincho E font literal ＭＳ 明朝, before 9, bold
  no_before  as real, no spacing before
  no_bold    as real, no <w:b/>
  known_sz24 E font ＭＳ 明朝 at sz 24 (the Arial-24 empties above them), before 9
  text_e     E paragraphs carry text 「あ」 (control: text vs empty)
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/emptybefore_grid'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
      '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>')
FONTS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
         '<w:font w:name="ＭＳ明朝,Bold"><w:altName w:val="ＭＳ 明朝"/><w:panose1 w:val="00000000000000000000"/><w:charset w:val="80"/><w:family w:val="auto"/><w:notTrueType/><w:pitch w:val="default"/></w:font>'
         '<w:font w:name="ＭＳ 明朝"><w:altName w:val="MS Mincho"/><w:panose1 w:val="02020609040205080304"/><w:charset w:val="80"/><w:family w:val="roman"/><w:pitch w:val="fixed"/></w:font>'
         '<w:font w:name="Century"><w:panose1 w:val="02040604050505020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
         '</w:fonts>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsiaTheme="minorEastAsia" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
          '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>'
          '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:eastAsia="ja-JP"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Body"><w:name w:val="Body"/><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="Courier New" w:cs="Times New Roman"/><w:lang w:eastAsia="en-US"/></w:rPr></w:style>'
          '</w:styles>')
SECT = ('<w:sectPr><w:pgSz w:w="11907" w:h="16840" w:code="9"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/><w:docGrid w:type="linesAndChars" w:linePitch="387"/></w:sectPr>')
SECT_LINES = SECT.replace('w:type="linesAndChars"', 'w:type="lines"')
STYLES_JA = STYLES.replace('w:eastAsia="en-US" w:bidi', 'w:eastAsia="ja-JP" w:bidi')
SECT_CS = SECT.replace('w:linePitch="387"/>', 'w:linePitch="387" w:charSpace="6144"/>')
SECT_CS_JA = SECT_CS


def para(text, rpr, before=True):
    sp = '<w:spacing w:before="180"/>' if before else ''
    ppr = f'<w:pPr><w:pStyle w:val="Body"/>{sp}<w:jc w:val="center"/><w:rPr>{rpr}</w:rPr></w:pPr>'
    run = f'<w:r><w:rPr>{rpr}</w:rPr><w:t>{text}</w:t></w:r>' if text else ''
    return f'<w:p>{ppr}{run}</w:p>'


A_RPR = '<w:rFonts w:ascii="Century" w:hAnsi="Century" w:cs="Century"/><w:sz w:val="37"/><w:szCs w:val="37"/><w:lang w:eastAsia="ja-JP"/>'
B_RPR = '<w:rFonts w:ascii="Century" w:hAnsi="Century"/><w:b/><w:sz w:val="36"/><w:lang w:eastAsia="ja-JP"/>'
E_REAL = '<w:rFonts w:ascii="ＭＳ明朝,Bold" w:eastAsia="ＭＳ明朝,Bold" w:hAnsi="Times New Roman" w:cs="ＭＳ明朝,Bold"/><w:b/><w:bCs/><w:sz w:val="37"/><w:szCs w:val="37"/>'
E_LIT = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="Times New Roman" w:cs="ＭＳ 明朝"/><w:b/><w:bCs/><w:sz w:val="37"/><w:szCs w:val="37"/>'
E_NOBOLD = '<w:rFonts w:ascii="ＭＳ明朝,Bold" w:eastAsia="ＭＳ明朝,Bold" w:hAnsi="Times New Roman" w:cs="ＭＳ明朝,Bold"/><w:sz w:val="37"/><w:szCs w:val="37"/>'
E_SZ24 = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="Times New Roman" w:cs="ＭＳ 明朝"/><w:b/><w:sz w:val="48"/><w:szCs w:val="48"/>'


def body(e_rpr, before=True, e_text=''):
    return (para('公表版1.0', A_RPR) + para(e_text, e_rpr, before) + para(e_text, e_rpr, before) + para(e_text, e_rpr, before)
            + para('2018年　6月', B_RPR, False) + para('', e_rpr, before))


arms = {
    'real': body(E_REAL),
    'lit_mincho': body(E_LIT),
    'no_before': body(E_REAL, before=False),
    'no_bold': body(E_NOBOLD),
    'known_sz24': body(E_SZ24),
    'text_e': body(E_REAL, e_text='あ'),
    'grid_lines': (body(E_REAL), SECT_LINES, None),
    'lang_ja': (body(E_REAL), None, STYLES_JA),
    'lines_ja': (body(E_REAL), SECT_LINES, STYLES_JA),
    'lac_cs': (body(E_REAL), SECT_CS, None),
    'lac_cs_ja': (body(E_REAL), SECT_CS, STYLES_JA),
    'lac_cs_text': (body(E_REAL, e_text='あ'), SECT_CS, STYLES_JA),
}


def document(b, sect=None):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{b}{sect or SECT}</w:body></w:document>'



import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try:
    app.Visible = False
except Exception:
    pass
try:
    for name, arm in arms.items():
        b, sect_v, styles_v = arm if isinstance(arm, tuple) else (arm, None, None)
        at = OUT / f"{name}.docx"
        with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", DOC_RELS); z.writestr("word/styles.xml", styles_v or STYLES)
            z.writestr("word/fontTable.xml", FONTS); z.writestr("word/document.xml", document(b, sect_v))
        d = app.Documents.Open(str(at.resolve()), False, True)
        try:
            n = d.Paragraphs.Count
            ys = [round(d.Range(d.Paragraphs(i).Range.Start, d.Paragraphs(i).Range.Start).Information(6), 2) for i in range(1, n + 1)]
            f = d.Paragraphs(2).Range.Font
            fn = (f.Name, f.NameFarEast)
        finally:
            d.Close(False)
        with tempfile.TemporaryDirectory() as t:
            dump = Path(t) / 'l.json'
            subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
            dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
        oy = sorted({round(el['y'], 2) for el in dd['pages'][0]['elements'] if el.get('type') == 'text'})
        wd = [round(b2 - a2, 2) for a2, b2 in zip(ys, ys[1:])]
        od = [round(b2 - a2, 2) for a2, b2 in zip(oy, oy[1:])]
        print(f"=== {name:11s} WORD y {ys} d {wd} font {fn}\n                OXI  y {oy} d {od}")
finally:
    app.Quit()
