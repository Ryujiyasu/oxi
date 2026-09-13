# -*- coding: utf-8 -*-
"""Which face prices U+221A (square root) on a Latin run under an empty-ea theme?

educational__003299ba4bee5126 (blindD50): rPrDefault eastAsiaTheme=minorEastAsia,
lang eastAsia ja-JP, theme minor <a:ea typeface=""/> with Jpan = 游明朝. Table
cells hold a lone '√' in Arial 10 (no w:hint). Word rows are 13.5pt, i.e. the
√ is priced like Arial 10; Oxi routes U+221A to the eastAsia face, which the
S327 stand-in "MS Mincho" hid and S1397's 游明朝 (line 16.7 at 10pt) exposes.

Arms: three paragraphs A1 / X / A2 in Arial 10 (A = 'Aa'), X varies:
  A_sqrt         '√' Arial 10, no hint
  B_sqrt_hint    '√' with w:hint="eastAsia"
  C_sqrt_ea_lit  '√' with explicit eastAsia="游明朝"
  D_kanji        '漢' Arial 10 (control: a real eastAsia char takes the Jpan face)
  E_sqrt_enUS    '√' with lang eastAsia=en-US on the run
  F_arrow        '→' (U+2192) no hint
Readout: Information(6) of A1/X/A2 -> X's line = y(A2)-y(X); Font.Name/NameFarEast of X.
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/sqrt_ea'); OUT.mkdir(parents=True, exist_ok=True)
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
         '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/></a:majorFont>'
         '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游明朝"/></a:minorFont></a:fontScheme>'
         '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
         '</a:themeElements></a:theme>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-GB" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '</w:styles>')
SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
AR = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'


def p(t, rf=AR, extra=''):
    rpr = f'<w:rPr>{rf}{extra}<w:sz w:val="20"/></w:rPr>'
    return f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>{rpr}</w:pPr><w:r>{rpr}<w:t>{t}</w:t></w:r></w:p>'


arms = {
    'A_sqrt': p('√'),
    'B_sqrt_hint': p('√', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:hint="eastAsia"/>'),
    'C_sqrt_ea_lit': p('√', '<w:rFonts w:ascii="Arial" w:eastAsia="游明朝" w:hAnsi="Arial" w:cs="Arial"/>'),
    'D_kanji': p('漢'),
    'E_sqrt_enUS': p('√', AR, '<w:lang w:eastAsia="en-US"/>'),
    'F_arrow': p('→'),
}


def document(x):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{p("Aa")}{x}{p("Aa")}{SECT}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try:
    app.Visible = False
except Exception:
    pass
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
            r = d.Paragraphs(2).Range
            ch = d.Range(r.Start, r.Start + 1)
            fn = (ch.Font.Name, ch.Font.NameFarEast, ch.Font.NameAscii)
        finally:
            d.Close(False)
        with tempfile.TemporaryDirectory() as t:
            dump = Path(t) / 'l.json'
            subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
            dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
        oy = [round(el['y'], 2) for el in dd['pages'][0]['elements'] if el.get('type') == 'text' and el.get('text', '').strip()]
        xl = round(oy[2] - oy[1], 2) if len(oy) > 2 else None
        print(f"=== {name}: WORD A1/X/A2 {ys}  X line {round(ys[2]-ys[1],2)} (A1 line {round(ys[1]-ys[0],2)})  font {fn}   OXI {oy[:3]} X line {xl}")
finally:
    app.Quit()
