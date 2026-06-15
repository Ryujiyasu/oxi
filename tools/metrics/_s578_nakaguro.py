# -*- coding: utf-8 -*-
"""S578 minimal repro — is ・ (nakaguro) compressed at RENDER on demand?

Build two jc=both / compress_punctuation / compat15 / MS Mincho 12pt docs:
  loose : short line, big right margin -> line not full -> ・ should stay ~12.0
  over  : long line, tight margin     -> line overflows -> ・ should compress

Measure ・ advance in Word (PDF render-truth = gold; COM Info(5) = cross-check)
and in Oxi (--dump-layout) with S578 OFF vs ON. Expectation:
  Word: loose ~12.0, over <12.0 (demand-driven).
  Oxi OFF: ・ stuck at 12.0 even when over (the bug).
  Oxi ON : ・ compresses on the over line, matching Word.
"""
import os, sys, zipfile, subprocess, json
sys.stdout.reconfigure(encoding='utf-8')
import fitz

OUT = os.path.abspath('tools/golden-test/repros/s578_nakaguro')
os.makedirs(OUT, exist_ok=True)
REPO = r'c:\Users\ryuji\oxi-main'
GDI = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
          '<w:kern w:val="2"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')

def build(docx, text, right_mar):
    body = ('<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="%d" w:bottom="1134" w:left="1304"/></w:sectPr>' % right_mar)
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>') % (body, sect)
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)

TAIL = u'あいうえおかきくけこさしすせそ'
# ・ after kanji, early on the line.
NAKA = u'・'  # ・
CASES = {
    'naka_loose': (u'国国' + NAKA + u'国国国', 6000),
    'naka_over':  (u'国' + NAKA + u'国' * 36 + TAIL, 1500),
}

def word_pdf_adv(docx, ch):
    pdf = docx.replace('.docx', '.pdf')
    import win32com.client as w32
    w = w32.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
        d.Close(False)
    finally:
        w.Quit()
    doc = fitz.open(pdf)
    advs = []
    for page in doc:
        for blk in page.get_text("rawdict").get("blocks", []):
            for ln in blk.get("lines", []):
                chars = []
                for sp in ln.get("spans", []):
                    for c in sp.get("chars", []):
                        chars.append((c["c"], c["bbox"][0]))
                for i, (c, x0) in enumerate(chars):
                    if c == ch and i + 1 < len(chars):
                        advs.append(round(chars[i+1][1] - x0, 2))
    return advs

def oxi_adv(docx, ch, disable):
    env = dict(os.environ)
    if disable: env['OXI_S578_DISABLE'] = '1'
    else: env.pop('OXI_S578_DISABLE', None)
    pref = docx.replace('.docx', '_oxi')
    dj = docx.replace('.docx', '_oxi.json')
    subprocess.run([GDI, os.path.abspath(docx), pref, '96', '--dump-layout=' + dj],
                   capture_output=True, env=env)
    j = json.load(open(dj, encoding='utf-8'))
    from collections import defaultdict
    advs = []
    for pg in j['pages']:
        rows = defaultdict(list)
        for el in pg['elements']:
            if el['type'] == 'text' and el['text']:
                rows[round(el['y'], 1)].append(el)
        for y in rows:
            els = sorted(rows[y], key=lambda e: e['x'])
            for e in els:
                if e['text'] == ch:
                    advs.append(round(e['w'], 2))
    return advs

print("=== S578 ・ (nakaguro) demand-compression repro (MS Mincho 12pt jc=both c15) ===")
for name, (text, rm) in CASES.items():
    docx = os.path.join(OUT, name + '.docx')
    build(docx, text, rm)
    wadv = word_pdf_adv(docx, NAKA)
    oadv_off = oxi_adv(docx, NAKA, True)
    oadv_on = oxi_adv(docx, NAKA, False)
    print(f"\n{name}: (Word PDF / Oxi OFF / Oxi ON)  natural=12.0")
    print(f"  Word ・ adv = {wadv}")
    print(f"  Oxi  OFF    = {oadv_off}")
    print(f"  Oxi  ON     = {oadv_on}")
