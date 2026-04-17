"""Test what char class before '（' triggers compression.

Fixed settings: cSC + compat15.
Fixed font/size: MS Gothic 12pt.
Vary: char BEFORE '（' across categories.
"""
import os, sys, time, json, zipfile
import win32com.client

TMP = os.path.abspath("pipeline_data/_ctx_tmp")
os.makedirs(TMP, exist_ok=True)

# Each test: (label, char_before, char_after)
TESTS = [
    ("kanji_before",       "漢", "字"),
    ("hiragana_before",    "あ", "字"),
    ("katakana_before",    "ア", "字"),
    ("halfdigit_before",   "1", "字"),
    ("halfletter_before",  "a", "字"),
    ("fullwidth_A_before", "Ａ", "字"),
    ("fullwidth_1_before", "１", "字"),
    ("kanji_after",        "漢", "あ"),   # kanji before, hiragana after
    ("hiragana_both",      "あ", "い"),
    ("kanji_both",         "漢", "字"),  # dup to confirm
    ("yak_before",         "」", "字"),   # yakumono before '（'
]

TEMPLATE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:linePitch="360"/></w:sectPr>
</w:body></w:document>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>'''

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''

def make_docx(out, text):
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', TEMPLATE.format(text=text))
        z.writestr('word/settings.xml', SETTINGS)

def measure(path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
        p = doc.Paragraphs(1)
        rng = p.Range
        chars = []
        for ci in range(1, rng.Characters.Count + 1):
            c = rng.Characters(ci)
            try:
                x = c.Information(5); y = c.Information(6)
                chars.append((c.Text, x, y))
            except: pass
        # Find '（' and its advance
        for i, (ch, x, y) in enumerate(chars):
            if ch == '（' and i + 1 < len(chars):
                nxt = chars[i+1]
                if abs(y - nxt[2]) < 2:
                    doc.Close(False)
                    return round(nxt[1] - x, 2)
        doc.Close(False)
        return None
    finally:
        word.Quit()

def main():
    print(f"{'test':<25} {'text':<10} {'（_adv':>8}")
    print('-' * 50)
    for label, before, after in TESTS:
        text = f"{before}（{after}"
        out = os.path.join(TMP, f"{label}.docx")
        make_docx(out, text)
        try:
            adv = measure(out)
            print(f"{label:<25} {text:<10} {adv if adv is not None else 'None':>8}")
        except Exception as e:
            print(f"{label:<25} ERROR: {e}")

main()
