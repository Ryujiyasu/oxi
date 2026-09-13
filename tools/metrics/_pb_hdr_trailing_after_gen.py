# -*- coding: utf-8 -*-
"""Does the LAST header paragraph's space-after push the body?

policies__0066a548f46e098c: two Lato-14 header paragraphs (line=259, after
from docDefaults 160 only). Word PDF page 3: header lines 35.34 / 53.58 /
79.64 / 97.72 (the 8pt after sits BETWEEN the paragraphs) and the body table's
top border at 119.42 = 97.72 + 18.13 + ~3.5 -- no trailing 8pt. S813
(uklocalspending, Normal before/after=240) and S1392's witness (educational,
direct after=240 + contextualSpacing) both COUNT the trailing after.

Arms (Calibri 11 body, header distance 720 = 36pt, top margin 36pt so the
header always binds; two one-line Arial-12 header paragraphs):
  A_docdefault_after   pPrDefault after=240, no style / direct spacing
  B_style_after        Normal style after=240
  C_direct_after       direct after=240 on both header paragraphs
  D_direct_last_only   direct after=240 on the LAST header paragraph only
  E_none               no after anywhere (control)
Readout: Word COM Information(6) of body paragraph 1 (= header bottom); Oxi.
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/hdr_trailing_after'); OUT.mkdir(parents=True, exist_ok=True)
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
def styles(ppr_default_after, normal_after):
    pd = f'<w:pPrDefault><w:pPr><w:spacing w:after="{ppr_default_after}" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    na = f'<w:pPr><w:spacing w:after="{normal_after}"/></w:pPr>' if normal_after is not None else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="MS Mincho" w:hAnsi="Calibri" w:cs="Times New Roman"/>'
            '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>' + pd + '</w:docDefaults>'
            f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>{na}</w:style>'
            '</w:styles>')
def hp(t, direct_after):
    sp = f'<w:spacing w:after="{direct_after}"/>' if direct_after is not None else ''
    rpr = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr>'
    return f'<w:p><w:pPr>{sp}{rpr}</w:pPr><w:r>{rpr}<w:t>{t}</w:t></w:r></w:p>'
def header(a1, a2):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{hp("Header line one", a1)}{hp("Header line two", a2)}</w:hdr>'
def document():
    body = ''.join(f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Body line {i}</w:t></w:r></w:p>' for i in range(1, 4))
    sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="720" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'
arms = {
    "A_docdefault_after": (240, None, None, None),
    "B_style_after": (0, 240, None, None),
    "C_direct_after": (0, None, 240, 240),
    "D_direct_last_only": (0, None, None, 240),
    "E_none": (0, None, None, None),
}
import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try: app.Visible = False
except Exception: pass
try:
  for name, (pda, na, a1, a2) in arms.items():
    at = OUT / f"{name}.docx"
    with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS); z.writestr("word/styles.xml", styles(pda, na))
        z.writestr("word/header1.xml", header(a1, a2)); z.writestr("word/document.xml", document())
    d = app.Documents.Open(str(at.resolve()), False, True)
    try:
        r = d.Paragraphs(1).Range; c = d.Range(r.Start, r.Start); body_y = round(c.Information(6), 2)
        hr = d.Sections(1).Headers(1).Range
        hy = [round(d.Range(hr.Paragraphs(i).Range.Start, hr.Paragraphs(i).Range.Start).Information(6), 2) for i in range(1, hr.Paragraphs.Count + 1)]
    finally: d.Close(False)
    with tempfile.TemporaryDirectory() as t:
        dump = Path(t) / 'l.json'
        subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
        dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
    oxi = sorted({round(el['y'], 2) for el in dd['pages'][0]['elements'] if el.get('type') == 'text' and 'Body' in el.get('text', '')})
    print(f"=== {name}: WORD header lines {hy} body {body_y}   OXI body {oxi[:1]}")
finally: app.Quit()
