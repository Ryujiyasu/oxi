# -*- coding: utf-8 -*-
"""A JAPANESE header holding ONE empty paragraph: does it push the body?

forms__009bbe8a5704589a (jablindC50): header1 = one empty paragraph (mark
ＭＳ 明朝 sz=22, kinsoku/autoSpace off), header distance 851 = 42.55, top
margin 851, lines grid 360. Word COM: first body line at 58.5 (= 42.55 +
15.95); Oxi: 42.55 (S843 ink-free header -> 0; S1064's one-bare-paragraph
collapse and its n>=2 reservation are Latin-scoped).

Arms (docDefaults Century/ＭＳ 明朝 10.5, header distance 851, top margin 851):
  A_one_sz22_grid     one empty header paragraph, mark sz=22, lines grid 360
  B_one_nosz_grid     mark without sz
  C_one_sz22_nogrid   no docGrid
  D_two_sz22_grid     two empty header paragraphs
  E_one_sz22_text     the same mark but WITH text ("見出し") -- control
  F_one_sz21_grid     mark sz=21 (10.5pt)
Readout: Word COM Information(6) of body paragraph 1 minus 42.55; Oxi.
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/jphdr_empty'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
      '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
          '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="a3"><w:name w:val="header"/><w:basedOn w:val="a"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4252"/><w:tab w:val="right" w:pos="8504"/></w:tabs><w:snapToGrid w:val="0"/></w:pPr></w:style>'
          '</w:styles>')
def hp(sz, text=''):
    s = f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>' if sz else ''
    rpr = f'<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>{s}</w:rPr>'
    run = f'<w:r>{rpr}<w:t>{text}</w:t></w:r>' if text else ''
    return f'<w:p><w:pPr><w:pStyle w:val="a3"/><w:kinsoku w:val="0"/><w:overflowPunct w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:jc w:val="left"/>{rpr}</w:pPr>{run}</w:p>'
def header(paras):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{"".join(paras)}</w:hdr>'
def document(grid):
    body = ''.join(f'<w:p><w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>本文{i}行目</w:t></w:r></w:p>' for i in range(1, 4))
    g = '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else ''
    sect = (f'<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="851" w:right="851" w:bottom="567" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>{g}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'
arms = {
    "A_one_sz22_grid": ([hp(22)], True),
    "B_one_nosz_grid": ([hp(None)], True),
    "C_one_sz22_nogrid": ([hp(22)], False),
    "D_two_sz22_grid": ([hp(22), hp(22)], True),
    "E_one_sz22_text": ([hp(22, '見出し')], True),
    "F_one_sz21_grid": ([hp(21)], True),
}
import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try: app.Visible = False
except Exception: pass
try:
  for name, (paras, grid) in arms.items():
    at = OUT / f"{name}.docx"
    with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/header1.xml", header(paras)); z.writestr("word/document.xml", document(grid))
    d = app.Documents.Open(str(at.resolve()), False, True)
    try:
        r = d.Paragraphs(1).Range; c = d.Range(r.Start, r.Start); body_y = round(c.Information(6), 2)
    finally: d.Close(False)
    with tempfile.TemporaryDirectory() as t:
        dump = Path(t) / 'l.json'
        subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
        dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
    oxi = sorted({round(el['y'], 2) for el in dd['pages'][0]['elements'] if el.get('type') == 'text' and '本文' in el.get('text', '')})
    print(f"=== {name}: WORD body {body_y} (-42.55 = {round(body_y - 42.55, 2)})   OXI body {oxi[:1]}")
finally: app.Quit()
