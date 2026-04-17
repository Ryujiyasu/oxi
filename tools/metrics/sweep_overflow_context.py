"""Test if yakumono compression is overflow-triggered.

Fixed: cSC+compat15, MS Gothic 12pt.
Vary: text length to force wrap or not.
Measure '（' advance at different positions.
"""
import os, sys, time, json, zipfile
import win32com.client

TMP = os.path.abspath("pipeline_data/_ctx_tmp")
os.makedirs(TMP, exist_ok=True)

# Each test: (label, text)
TESTS = [
    # Short: no overflow expected
    ("short_kanji_open",     "漢（字あい"),
    # Long: multi-line, should overflow
    ("long_with_open",       "漢字あいうえおかきくけこさしすせそたちつてとなにぬねの漢字利用規約（第1.0版）はばひふへほまみむめも"),
    # d77a-like pattern
    ("d77a_like",            "「公共データ利用規約（第1.0版）」の前身である「政府標準利用規約」は、各府省ウェブサイトの利用ルー"),
    # Very long repetition
    ("wrap_many",            "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん漢字用規（第）版"),
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

CT = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'

def make_docx(out, text):
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', TEMPLATE.format(text=text))
        z.writestr('word/settings.xml', SETTINGS)

def measure_all_yak(path):
    """Return list of (char, x, y, advance) for all yakumono."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    YAK = '（）「」、。'
    try:
        doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
        p = doc.Paragraphs(1)
        rng = p.Range
        chars = []
        for ci in range(1, rng.Characters.Count + 1):
            c = rng.Characters(ci)
            try:
                x = c.Information(5); y = c.Information(6)
                chars.append((c.Text, round(x, 2), round(y, 2)))
            except: pass
        results = []
        for i, (ch, x, y) in enumerate(chars):
            if ch in YAK and i + 1 < len(chars):
                nxt = chars[i+1]
                if abs(y - nxt[2]) < 2:
                    results.append((ch, x, y, round(nxt[1] - x, 2)))
        doc.Close(False)
        return results
    finally:
        word.Quit()

def main():
    for label, text in TESTS:
        out = os.path.join(TMP, f"overflow_{label}.docx")
        make_docx(out, text)
        try:
            yaks = measure_all_yak(out)
            print(f"\n{label}: {text[:30]!r}...")
            for ch, x, y, adv in yaks:
                print(f"  '{ch}' x={x:>7} y={y:>6} advance={adv:>5}")
        except Exception as e:
            print(f"{label}: ERROR {e}")

main()
