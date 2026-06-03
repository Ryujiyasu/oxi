# -*- coding: utf-8 -*-
"""S492k — isolate Word's CJK<->digit autoSpace at fs=12, docGrid linesAndChars
(b837 context). b837 idx-? para showed 第(before digit 1) advance 15.75 in Word vs
Oxi 15.0 -> autoSpace 3.75 vs Oxi 3.0. Confirm on a minimal repro before any change
(autoSpace is Phase-1-sensitive). Patterns: 国1国1... (CJK<->ASCII digit). Measure
Word per-char advance; report the CJK-before-digit and digit-before-CJK advances.
cp932-safe: UTF-8 file, ASCII output.
"""
import zipfile, os
import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/autospace_digit')
os.makedirs(OUT, exist_ok=True)
DOCX = os.path.join(OUT, 'as_digit.docx')
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
# b837 settings: autoSpaceDE/DN are ON by default unless turned off; include nothing special
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')
texts = ['国1国2国3国4国5国6国7国8国9国0国',  # CJK<->ascii-digit alternating
         '第1第2第3第4第5第6第7第8第9第0第']  # 第 (the actual b837 char) <-> digit
body = ''.join(
    ('<w:p><w:pPr><w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝"/><w:sz w:val="24"/></w:rPr></w:pPr>'
     '<w:r><w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝"/><w:sz w:val="24"/></w:rPr>'
     '<w:t xml:space="preserve">%s</w:t></w:r></w:p>') % t for t in texts)
doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>%s'
       '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/>'
       '<w:docGrid w:type="linesAndChars" w:linePitch="360"/></w:sectPr></w:body></w:document>') % body
with zipfile.ZipFile(DOCX, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
    z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/document.xml', doc)
    z.writestr('word/styles.xml', STYLES); z.writestr('word/settings.xml', SETTINGS)

WD_HPOS = 5
word = w32.DispatchEx('Word.Application'); word.Visible = False
try:
    wdoc = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    try:
        for pi, p in enumerate(wdoc.Paragraphs):
            rng = p.Range; txt = rng.Text; start = rng.Start
            seq = []
            for i in range(min(len(txt), 24)):
                ch = txt[i]
                if ch in ('\r', '\n', '\x07'):
                    continue
                x = wdoc.Range(start + i, start + i).Information(WD_HPOS)
                seq.append((ch, x))
            print("para %d:" % pi)
            for j in range(len(seq) - 1):
                ch = seq[j][0]; adv = round(seq[j + 1][1] - seq[j][1], 2)
                kind = 'CJK' if ord(ch) > 0x2000 else 'ascii'
                nxt = seq[j + 1][0]
                nkind = 'CJK' if ord(nxt) > 0x2000 else 'ascii'
                note = ''
                if kind == 'CJK' and nkind == 'ascii':
                    note = ' CJK->digit autoSpace=%.2f' % (adv - 12.0)
                elif kind == 'ascii' and nkind == 'CJK':
                    note = ' digit->CJK (adv incl autoSpace on digit)'
                print("   %-7s adv=%.2f%s" % (hex(ord(ch)), adv, note))
    finally:
        wdoc.Close(False)
finally:
    word.Quit()
