# -*- coding: utf-8 -*-
"""S492e — minimal repro: does linesAndChars respect paragraph indent in the line
break? Word should fit (charsLine - indent_cells) chars on an indented line; the
b837 evidence says Oxi fits the FULL charsLine (overflows the right boundary by the
indent width). Generate a docGrid=linesAndChars doc with 3 paras (no indent /
leftChars=200 / leftChars=400) of long CJK, measure Word chars-per-line via COM,
and emit ASCII results to a file (cp932 console can't show Japanese; heredocs mangle).
"""
import zipfile, os, json
import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/lac_indent')
os.makedirs(OUT, exist_ok=True)
DOCX = os.path.join(OUT, 'lac_indent.docx')

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
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')

KANJI = '国' * 60  # 60 x 国, pure CJK ideograph (no punct, no katakana -> isolate the indent)


def para(ind_chars, ind_tw):
    pind = ''
    if ind_chars:
        pind = '<w:ind w:leftChars="%d" w:left="%d"/>' % (ind_chars, ind_tw)
    return ('<w:p><w:pPr>%s<w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝"/><w:sz w:val="24"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝"/><w:sz w:val="24"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>') % (pind, KANJI)


body = para(0, 0) + para(200, 480) + para(400, 960)  # 0 / 2-char(24pt) / 4-char(48pt) indent
doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>%s'
       '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/>'
       '<w:docGrid w:type="linesAndChars" w:linePitch="360"/></w:sectPr></w:body></w:document>') % body
with zipfile.ZipFile(DOCX, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
    z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/document.xml', doc)
    z.writestr('word/styles.xml', STYLES); z.writestr('word/settings.xml', SETTINGS)

# --- Word L1 char count per paragraph ---
WD_VPOS = 6
word = w32.DispatchEx('Word.Application'); word.Visible = False
res = []
try:
    wdoc = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    try:
        for pi, p in enumerate(wdoc.Paragraphs):
            rng = p.Range; txt = rng.Text; start = rng.Start
            y0 = doc and wdoc.Range(start, start).Information(WD_VPOS)
            n = 0
            for i in range(len(txt)):
                if txt[i] in ('\r', '\n', '\x07'):
                    continue
                if wdoc.Range(start + i, start + i).Information(WD_VPOS) > y0 + 2:
                    break
                n += 1
            res.append({'para': pi, 'left_indent_pt': round(p.LeftIndent, 2), 'word_L1': n})
    finally:
        wdoc.Close(False)
finally:
    word.Quit()

json.dump(res, open('c:/tmp/lac_indent_word.json', 'w'), ensure_ascii=False, indent=1)
print("Word L1 per para (indent_pt -> L1 chars):")
for r in res:
    print("  para%d indent=%.1fpt  word_L1=%d" % (r['para'], r['left_indent_pt'], r['word_L1']))
print("\nExpected if Word respects indent: L1 drops ~2 per +24pt(2-char) indent.")
print("docx written:", DOCX)
