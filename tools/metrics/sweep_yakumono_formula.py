"""Sweep fs × font × cSC+compat15 to find yakumono compression formula.

Creates minimal docx with:
- characterSpacingControl=compressPunctuation
- compatibilityMode=15
- Single paragraph with yakumono: "あい（う）、え「お」か。きく"
- Font size varies; font family MS Gothic and MS Mincho

Measures '（', '、', '。' advance per fs per font.
"""
import os, sys, time, json, zipfile, tempfile
import win32com.client

TMP = os.path.abspath("pipeline_data/_sweep_tmp")
os.makedirs(TMP, exist_ok=True)

FONTS = [
    ("ＭＳ ゴシック", "MSGothic"),
    ("ＭＳ 明朝", "MSMincho"),
]
SIZES = [9.0, 10.5, 11.0, 12.0, 14.0, 16.0]

YAK_CHARS = ['（', '）', '、', '。', '「', '」']

TEMPLATE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz_half}"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz_half}"/></w:rPr><w:t xml:space="preserve">あい（う）、え「お」か。きく</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:linePitch="360"/></w:sectPr>
</w:body></w:document>'''

SETTINGS_TEMPLATE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>'''

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

def make_docx(out_path, font, sz_pt):
    sz_half = int(sz_pt * 2)
    doc_xml = TEMPLATE.format(font=font, sz_half=sz_half)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc_xml)
        z.writestr('word/settings.xml', SETTINGS_TEMPLATE)

def measure_chars(docx_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True); time.sleep(0.3)
        p = doc.Paragraphs(1)
        rng = p.Range
        n = rng.Characters.Count
        chars = []
        for ci in range(1, n + 1):
            c = rng.Characters(ci)
            try:
                x = c.Information(5); y = c.Information(6)
                chars.append({"ch": c.Text, "x": round(x, 2), "y": round(y, 2)})
            except: pass
        # Compute advances for yakumono chars
        advances = {}
        for i, ch in enumerate(chars):
            if ch["ch"] in YAK_CHARS and i + 1 < len(chars):
                nxt = chars[i + 1]
                if abs(ch["y"] - nxt["y"]) < 2:
                    adv = round(nxt["x"] - ch["x"], 2)
                    advances[ch["ch"]] = adv
        doc.Close(False)
        return advances
    finally:
        word.Quit()

def main():
    results = {}
    for font_name, font_key in FONTS:
        for sz in SIZES:
            key = f"{font_key}_{sz}"
            out = os.path.join(TMP, f"sweep_{key}.docx")
            try:
                make_docx(out, font_name, sz)
                adv = measure_chars(out)
                results[key] = {"font": font_key, "size": sz, "advances": adv}
                print(f"  {key:<25} {adv}")
            except Exception as e:
                results[key] = {"font": font_key, "size": sz, "error": str(e)}
                print(f"  {key:<25} ERROR {e}")

    with open("pipeline_data/yakumono_sweep.json", "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print("\nSaved: pipeline_data/yakumono_sweep.json")

    print("\n=== Derived widths ===")
    print(f"{'font':<10} {'fs':>5} {'（':>5} {'）':>5} {'、':>5} {'。':>5} {'「':>5} {'」':>5}")
    for key, r in results.items():
        if "advances" in r:
            adv = r["advances"]
            print(f"{r['font']:<10} {r['size']:>5} "
                  + " ".join(f"{adv.get(c, '-'):>5}" for c in YAK_CHARS))

main()
