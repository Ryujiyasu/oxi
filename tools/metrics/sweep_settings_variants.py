"""Find the remaining trigger for '（' compression.

Fixed: cSC+compat15, MS Gothic 12pt, d77a-like text.
Vary: docGrid.type, styles.xml presence, fontTable.xml presence, etc.
"""
import os, sys, time, json, zipfile
import win32com.client

TMP = os.path.abspath("pipeline_data/_ctx_tmp")
os.makedirs(TMP, exist_ok=True)

TEXT = "「公共データ利用規約（第1.0版）」の前身である「政府標準利用規約」は、各府省ウェブサイトの利用ルー"

DOC_BASE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/>{extra_sec}</w:sectPr>
</w:body></w:document>'''

SETTINGS_BASE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>
{extra_compat}
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>'''

STYLES_MINIMAL = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>'''

FONTTABLE_MINIMAL = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:font w:name="ＭＳ ゴシック"><w:panose1 w:val="020B0609000101010101"/><w:charset w:val="80"/><w:family w:val="modern"/><w:pitch w:val="fixed"/></w:font>
<w:font w:name="ＭＳ 明朝"><w:panose1 w:val="02020609040205080304"/><w:charset w:val="80"/><w:family w:val="roman"/><w:pitch w:val="fixed"/></w:font>
</w:fonts>'''

CT_BASE = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>{extra_override}</Types>'
RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>{extra_rel}</Relationships>'

TESTS = [
    ("base",                     {}),
    ("+docGrid_lines",           {"extra_sec": '<w:docGrid w:type="lines" w:linePitch="360"/>'}),
    ("+docGrid_linesAndChars",   {"extra_sec": '<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="0"/>'}),
    ("+styles",                  {"extra_files": [("word/styles.xml", STYLES_MINIMAL)], "override": '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>', "rel": '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'}),
    ("+fontTable",               {"extra_files": [("word/fontTable.xml", FONTTABLE_MINIMAL)], "override": '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>', "rel": '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>'}),
    ("+useFELayout_compat",      {"extra_compat": '<w:useFELayout/>'}),
    ("+balanceByte",             {"extra_compat": '<w:balanceSingleByteDoubleByteWidth/>'}),
    ("+spaceForUL",              {"extra_compat": '<w:spaceForUL/>'}),
    ("+all_d77a_compat",         {"extra_compat": '<w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/><w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/><w:useFELayout/>'}),
    ("+docGrid_lines_+compat_all+styles+font",  {
        "extra_sec": '<w:docGrid w:type="lines" w:linePitch="360"/>',
        "extra_compat": '<w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/><w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/><w:useFELayout/>',
        "extra_files": [("word/styles.xml", STYLES_MINIMAL), ("word/fontTable.xml", FONTTABLE_MINIMAL)],
        "override": '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>',
        "rel": '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>'
    }),
]

def make_docx(out, cfg):
    extra_sec = cfg.get("extra_sec", "")
    extra_compat = cfg.get("extra_compat", "")
    extra_files = cfg.get("extra_files", [])
    override = cfg.get("override", "")
    rel = cfg.get("rel", "")

    doc_xml = DOC_BASE.format(text=TEXT, extra_sec=extra_sec)
    settings_xml = SETTINGS_BASE.format(extra_compat=extra_compat)
    ct_xml = CT_BASE.format(extra_override=override)
    doc_rels = DOC_RELS.format(extra_rel=rel)

    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct_xml)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('word/document.xml', doc_xml)
        z.writestr('word/settings.xml', settings_xml)
        for name, content in extra_files:
            z.writestr(name, content)

def measure_open(path):
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
                chars.append((c.Text, round(x, 2), round(y, 2)))
            except: pass
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
    print(f"{'variant':<45} {'（_adv':>8}")
    print('-' * 58)
    for label, cfg in TESTS:
        out = os.path.join(TMP, f"var_{label.replace('+','_').replace('/','_')}.docx")
        try:
            make_docx(out, cfg)
            adv = measure_open(out)
            print(f"{label:<45} {adv if adv is not None else 'None':>8}")
        except Exception as e:
            print(f"{label:<45} ERROR: {e}")

main()
