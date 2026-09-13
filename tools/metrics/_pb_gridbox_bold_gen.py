# -*- coding: utf-8 -*-
"""Why does a 12pt bold ＭＳ ゴシック heading take TWO grid lines on a 16.45 grid?

policies__0820fb071316448b (jablindC50): docGrid lines, linePitch 329
(16.45pt); docDefaults sz=21 but the document's Normal-inherited size is 12pt
(drs Century 12). The heading '２ 患者発生時の患者、濃厚接触者への対応'
(majorEastAsia = ＭＳ ゴシック, bold, hanging indent, kern=0) sits in a
two-line grid box in Word -- COM: the empty line before it at 520.5, the
heading at 544.5 (+24), the next paragraph at 570 (+25.5): 16.45 + 32.9 laid
out as 7.5 above / 9 below the text (the em box centred in a 2-line box).
Oxi: ＭＳ ゴシック 12pt box 15.5 <= 16.45 -> one line.

Arms (docDefaults sz=24 to make 12pt the chain size, docGrid lines 329,
anchor A1 / X / A2; X varies):
  A_gothic_bold     rFonts ＭＳ ゴシック (literal), b
  B_gothic_plain    rFonts ＭＳ ゴシック, no b
  C_theme_major     rFonts majorEastAsia theme (theme major Jpan = ＭＳ ゴシック), b
  D_mincho_bold     rFonts ＭＳ 明朝 literal, b
  E_gothic_bold_11  as A at sz=22 (11pt)
  F_gothic_bold_105 as A at sz=21 (10.5pt)
Readout: COM Information(6) of A1 / X / A2 -> X's box = y(A2) - y(X).
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/gridbox_bold'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
      '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>')
THEME = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements>'
         '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
         '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ ゴシック"/></a:majorFont>'
         '<a:minorFont><a:latin typeface="Century"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ 明朝"/></a:minorFont></a:fontScheme>'
         '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
         '</a:themeElements></a:theme>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
          '<w:kern w:val="2"/><w:sz w:val="24"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '</w:styles>')
SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="851" w:right="851" w:bottom="851" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:docGrid w:type="lines" w:linePitch="329"/></w:sectPr>')
def p(t, rf, bold, sz=None):
    b = '<w:b/>' if bold else ''
    s = f'<w:sz w:val="{sz}"/>' if sz else ''
    rpr = f'<w:rPr>{rf}{b}<w:kern w:val="0"/>{s}<w:szCs w:val="21"/></w:rPr>'
    return f'<w:p><w:pPr><w:ind w:left="241" w:hangingChars="100" w:hanging="241"/>{rpr}</w:pPr><w:r>{rpr}<w:t>{t}</w:t></w:r></w:p>'
def anchor(t):
    rf = '<w:rFonts w:asciiTheme="minorEastAsia" w:hAnsiTheme="minorEastAsia" w:cs="ＭＳ Ｐゴシック"/>'
    return p(t, rf, False)
GOTH = '<w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:cs="ＭＳ Ｐゴシック"/>'
MAJ = '<w:rFonts w:asciiTheme="majorEastAsia" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorEastAsia" w:cs="ＭＳ Ｐゴシック"/>'
MIN = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:cs="ＭＳ Ｐゴシック"/>'
X = '２ 患者発生時の患者、濃厚接触者への対応'
arms = {
    "A_gothic_bold": p(X, GOTH, True),
    "B_gothic_plain": p(X, GOTH, False),
    "C_theme_major": p(X, MAJ, True),
    "D_mincho_bold": p(X, MIN, True),
    "E_gothic_bold_11": p(X, GOTH, True, 22),
    "F_gothic_bold_105": p(X, GOTH, True, 21),
}
def document(x):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{anchor("A1")}{x}{anchor("A2")}{SECT}</w:body></w:document>'
import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try: app.Visible = False
except Exception: pass
try:
  for name, x in arms.items():
    at = OUT / f"{name}.docx"
    with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/theme/theme1.xml", THEME); z.writestr("word/document.xml", document(x))
    d = app.Documents.Open(str(at.resolve()), False, True)
    try:
        ys = [round(d.Range(d.Paragraphs(i).Range.Start, d.Paragraphs(i).Range.Start).Information(6), 2) for i in (1, 2, 3)]
    finally: d.Close(False)
    with tempfile.TemporaryDirectory() as t:
        dump = Path(t) / 'l.json'
        subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
        dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
    oy = {}
    for el in dd['pages'][0]['elements']:
        if el.get('type') == 'text' and el.get('text', '').strip():
            k = el['text'][:2]
            oy.setdefault(k, round(el['y'], 2))
    print(f"=== {name}: WORD A1/X/A2 {ys}  X box {round(ys[2]-ys[1],2)} (A1->X {round(ys[1]-ys[0],2)})   OXI A1/X/A2 {oy.get('A1')}/{oy.get('２ ')}/{oy.get('A2')}")
finally: app.Quit()
