# -*- coding: utf-8 -*-
"""S492l — reproduce b837's EXACT indent structure to find why Oxi natural over-packs
some li=24 paras (fits 37, boundary-overflow) while the simpler lac_indent repro was
correct (35). b837's 本府 para has leftChars=200 left=480 firstLineChars=100
firstLine=240 + pStyle Web. Test pure-国 with this exact structure: does Oxi over-pack
the continuation lines vs Word? Isolates whether firstLine/pStyle breaks the break-time
available-width (indent applied to render x but not to the break). cp932-safe.
"""
import zipfile, os, subprocess, json
import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/b837indent')
os.makedirs(OUT, exist_ok=True)
DOCX = os.path.join(OUT, 'b837indent.docx')
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
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
# include a "Web" style (HTML preformatted-ish) like b837
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Web"><w:name w:val="Normal (Web)"/></w:style></w:styles>')

# variant A: jc=both, exact b837 indent, pStyle Web, pure 国
# variant B: same but jc=left
def para(jc):
    return ('<w:p><w:pPr><w:pStyle w:val="Web"/>'
            '<w:ind w:leftChars="200" w:left="480" w:firstLineChars="100" w:firstLine="240"/>'
            '<w:jc w:val="%s"/><w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝"/><w:sz w:val="24"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝"/><w:sz w:val="24"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>') % (jc, '国' * 80)

body = para('both') + para('left')
doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>%s'
       '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1021" w:right="1418" w:bottom="1021" w:left="1418"/>'
       '<w:docGrid w:type="linesAndChars" w:linePitch="360"/></w:sectPr></w:body></w:document>') % body
with zipfile.ZipFile(DOCX, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
    z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/document.xml', doc)
    z.writestr('word/styles.xml', STYLES); z.writestr('word/settings.xml', SETTINGS)

# Oxi render
BIN = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
subprocess.run([BIN, os.path.abspath(DOCX), 'c:/tmp/_bi', '--dump-layout=c:/tmp/_bi.json'],
               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
dd = json.load(open('c:/tmp/_bi.json', encoding='utf-8'))
print("boundary: content 453.5 from left margin 70.9 -> right boundary 524.4pt")
for pi in (0, 1):
    els = [e for pg in dd['pages'] for e in pg['elements'] if e['type'] == 'text' and e.get('para_idx') == pi]
    els.sort(key=lambda e: (round(e['y'], 1), e['x']))
    lines = {}
    for e in els:
        lines.setdefault(round(e['y'], 1), []).append(e)
    print("Oxi para%d (%s):" % (pi, 'both' if pi == 0 else 'left'))
    for y in sorted(lines)[:4]:
        ln = sorted(lines[y], key=lambda e: e['x'])
        print("   y=%.1f n=%d x0=%.1f xend=%.1f%s" % (y, len(ln), ln[0]['x'], ln[-1]['x'] + ln[-1]['w'],
              '  OVERFLOW' if ln[-1]['x'] + ln[-1]['w'] > 525.4 else ''))

# Word measure
WD_VPOS, WD_HPOS = 6, 5
word = w32.DispatchEx('Word.Application'); word.Visible = False
try:
    wdoc = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    try:
        for pi, p in enumerate(wdoc.Paragraphs):
            rng = p.Range; txt = rng.Text; start = rng.Start
            y0 = wdoc.Range(start, start).Information(WD_VPOS)
            counts = []; cur = 0; prev = y0
            for i in range(len(txt)):
                if txt[i] in ('\r', '\n', '\x07'):
                    continue
                y = wdoc.Range(start + i, start + i).Information(WD_VPOS)
                if y > prev + 2:
                    counts.append(cur); cur = 0; prev = y
                cur += 1
            counts.append(cur)
            print("Word para%d (li=%.0f fli=%.0f jc=%d): lines=%s" %
                  (pi, p.LeftIndent, p.FirstLineIndent, p.Alignment, counts[:5]))
    finally:
        wdoc.Close(False)
finally:
    word.Quit()
