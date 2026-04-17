"""Tier 1 additive: scratch + cSC + compat15 + kern + jc + docGrid type="lines".

If trigger → compression activates (fs=12 '（' → 10.5pt).
If not → trigger is elsewhere (move to Tier 2).
"""
import os, sys, time, zipfile
import win32com.client

TMP = os.path.abspath("pipeline_data/_additive_tmp")
os.makedirs(TMP, exist_ok=True)

def make_docx(out_path, docgrid_xml, fs=12):
    sz = int(fs * 2)
    doc = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:jc w:val="both"/><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:kern w:val="2"/><w:sz w:val="{sz}"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:kern w:val="2"/><w:sz w:val="{sz}"/></w:rPr><w:t xml:space="preserve">「公共データ利用規約（第1.0版）」の前身である「政府標準利用規約」は、各府省ウェブサイトの利用ルー</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/>{docgrid_xml}</w:sectPr>
</w:body></w:document>'''
    settings = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''
    ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
    rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    doc_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('word/document.xml', doc)
        z.writestr('word/settings.xml', settings)

def measure(path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
        p = doc.Paragraphs(1)
        rng = p.Range
        for ci in range(1, rng.Characters.Count + 1):
            c = rng.Characters(ci)
            if c.Text == '（':
                try:
                    x1 = c.Information(5); y1 = c.Information(6)
                    nxt = rng.Characters(ci + 1)
                    x2 = nxt.Information(5); y2 = nxt.Information(6)
                    if abs(y1 - y2) > 2: continue
                    doc.Close(False)
                    return round(x2 - x1, 2)
                except: pass
        doc.Close(False)
        return None
    finally:
        try: word.Quit()
        except: pass

# Tier 1 variants — test docGrid with and without type attribute
TESTS = [
    ("no_docGrid",           ''),
    ("docGrid_linePitch_only", '<w:docGrid w:linePitch="360"/>'),
    ("docGrid_type_lines",   '<w:docGrid w:type="lines" w:linePitch="360"/>'),
    ("docGrid_linesAndChars", '<w:docGrid w:type="linesAndChars" w:linePitch="360"/>'),
    ("docGrid_default",      '<w:docGrid w:type="default" w:linePitch="360"/>'),
]

print(f"{'variant':<30}  fs=12 '（' adv")
print('-' * 50)
for label, dg in TESTS:
    out = os.path.join(TMP, f"t1_{label}.docx")
    make_docx(out, dg, fs=12)
    try:
        adv = measure(out)
        marker = ''
        if adv is not None:
            if adv < 11.0: marker = ' **COMPRESSED (HIT!)**'
            elif adv < 11.8: marker = ' (partial)'
            else: marker = ' (no)'
        print(f"{label:<30}  {adv if adv else 'None':>6}{marker}")
    except Exception as e:
        print(f"{label:<30}  ERROR {e}")
    time.sleep(1)
