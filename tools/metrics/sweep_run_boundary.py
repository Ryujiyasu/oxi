"""Test if run boundary before '（' triggers compression.

Matches d77a's actual run structure: Run1 ends with kanji, Run2 starts with '（'.
"""
import os, sys, time, zipfile
import win32com.client

TMP = os.path.abspath("pipeline_data/_ctx_tmp")
os.makedirs(TMP, exist_ok=True)

def make_docx(out, runs_list):
    """runs_list: list of text strings, each becomes a separate <w:r>."""
    runs_xml = ""
    for txt in runs_list:
        runs_xml += f'<w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:hint="eastAsia"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{txt}</w:t></w:r>'
    doc = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="24"/></w:rPr></w:pPr>{runs_xml}</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>
</w:body></w:document>'''
    settings = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''
    ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
    rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    doc_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('word/document.xml', doc)
        z.writestr('word/settings.xml', settings)

def measure_yak(path):
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

TESTS = [
    ("single_run",     ["「公共データ利用規約（第1.0版）」の前身である"]),
    ("2run_kanji_open", ["「公共データ利用規約", "（第1.0版）」の前身である"]),
    ("3run_d77a_like", ["「公共データ利用規約", "（第1.0版）", "」の前身である"]),
    ("2run_split_elsewhere", ["「公共データ", "利用規約（第1.0版）」の前身である"]),
]

for label, runs in TESTS:
    out = os.path.join(TMP, f"runs_{label}.docx")
    make_docx(out, runs)
    try:
        yaks = measure_yak(out)
        print(f"\n{label}: {' | '.join(runs)}")
        for ch, x, y, adv in yaks:
            marker = ' **COMPRESSED**' if adv < 11.5 else ''
            print(f"  '{ch}' x={x:>7} advance={adv}{marker}")
    except Exception as e:
        print(f"{label}: ERROR {e}")
