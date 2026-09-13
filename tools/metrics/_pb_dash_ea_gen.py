# -*- coding: utf-8 -*-
"""A lone en dash / curly quote RUN in Latin text: which face prices its line?

S763b/S763c (db9ca): a run holding a bare quote, in a doc whose eastAsia lang
is ja-JP and whose eastAsia font is a real CJK face, draws in that face. Their
inputs were literal ＭＳ Ｐゴシック. educational__003299ba4bee5126 (EN, lang
eastAsia ja-JP, rPrDefault minorEastAsia against an empty-ea theme whose Jpan
face is 游明朝) holds «(Birks et al, 2008) –» + a run of just «–», and Word's
paragraph is 10 x 14.5 = the Calibri 11 x 1.08 line -- the dash line did NOT
grow to 游明朝's 19.9. So the eastAsia font's ORIGIN (theme vs literal) or its
identity is a discriminator the quote law never saw.

Arms (Calibri 11 body, line 1.0, three paragraphs A1 / X / A2; X = 'abc' + a
LONE run of the symbol + 'def', all three runs plain):
  theme_dash    rPrDefault eastAsiaTheme=minorEastAsia, theme ea "" Jpan 游明朝, X = –
  theme_quote   same, X = ’
  lit_yu_dash   rPrDefault eastAsia="游明朝" literal, X = –
  lit_yu_quote  same, X = ’
  lit_pg_dash   rPrDefault eastAsia="ＭＳ Ｐゴシック" literal, X = –
  lit_pg_quote  same, X = ’
  theme_hint    theme shape, X = – with w:hint="eastAsia"
  theme_kanji   theme shape, X = 漢 (control)
Readout: X line = y(A2) - y(X) by Information(6); Calibri 11 = 13.43 (COM 13.5),
游明朝 11 = 18.4, ＭＳ Ｐゴシック 11 = 12.65 (83/64 box 14.27 at grid).
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/dash_ea'); OUT.mkdir(parents=True, exist_ok=True)
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


def styles(ea_attr):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" {ea_attr} w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
            '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-GB" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
            '</w:styles>')


THEME_EA = 'w:eastAsiaTheme="minorEastAsia"'
LIT_YU = 'w:eastAsia="游明朝"'
LIT_PG = 'w:eastAsia="ＭＳ Ｐゴシック"'
SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
PPR = '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'


def para(runs):
    return '<w:p>' + PPR + ''.join(runs) + '</w:p>'


def run(t, rpr=''):
    return f'<w:r>{rpr}<w:t xml:space="preserve">{t}</w:t></w:r>'


def x(sym, rpr=''):
    return para([run('abc '), run(sym, rpr), run(' def')])


HINT = '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
arms = {
    'theme_dash': (THEME_EA, x('–')),
    'theme_quote': (THEME_EA, x('’')),
    'lit_yu_dash': (LIT_YU, x('–')),
    'lit_yu_quote': (LIT_YU, x('’')),
    'lit_pg_dash': (LIT_PG, x('–')),
    'lit_pg_quote': (LIT_PG, x('’')),
    'theme_hint': (THEME_EA, x('–', HINT)),
    'theme_kanji': (THEME_EA, x('漢')),
}


def document(xp):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{para([run("Aa")])}{xp}{para([run("Aa")])}{SECT}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try:
    app.Visible = False
except Exception:
    pass
try:
    for name, (ea, xp) in arms.items():
        at = OUT / f"{name}.docx"
        with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", DOC_RELS); z.writestr("word/styles.xml", styles(ea))
            z.writestr("word/theme/theme1.xml", THEME); z.writestr("word/document.xml", document(xp))
        d = app.Documents.Open(str(at.resolve()), False, True)
        try:
            ys = [round(d.Range(d.Paragraphs(i).Range.Start, d.Paragraphs(i).Range.Start).Information(6), 2) for i in (1, 2, 3)]
        finally:
            d.Close(False)
        with tempfile.TemporaryDirectory() as t:
            dump = Path(t) / 'l.json'
            subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
            dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
        oy = sorted({round(el['y'], 2) for el in dd['pages'][0]['elements'] if el.get('type') == 'text' and el.get('text', '').strip()})
        xl = round(oy[2] - oy[1], 2) if len(oy) > 2 else None
        print(f"=== {name:12s}: WORD X line {round(ys[2]-ys[1],2):6} (A1 {round(ys[1]-ys[0],2)})   OXI X line {xl} (A1 {round(oy[1]-oy[0],2) if len(oy)>1 else None})")
finally:
    app.Quit()
