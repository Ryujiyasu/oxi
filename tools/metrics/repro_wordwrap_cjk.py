# -*- coding: utf-8 -*-
"""V_JJ: Test Word's wordWrap=off behavior on CJK vs Latin text.

ECMA-376 §17.3.1.40: wordWrap controls LATIN word breaking only.
- wordWrap=true (default): Latin can break mid-word
- wordWrap=false: Latin breaks at whitespace only

CJK should always break at char boundaries regardless.

Test if Word actually does this, and if Oxi's gating CJK on word_wrap is wrong."""
import os, sys, time, json, zipfile, shutil, tempfile
import pythoncom, win32com.client as wc
import fitz

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES, RELS_ROOT, WORD_RELS

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/v_jj_wordwrap")
os.makedirs(OUT_DIR, exist_ok=True)

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:kern w:val="2"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:defaultTabStop w:val="840"/>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>
</w:settings>'''


def doc_xml(*, paragraphs):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{paragraphs}
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:space="425"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>'''


def para(*, ind_xml="", word_wrap_xml="", text):
    return f'''<w:p>
<w:pPr>
{ind_xml}
{word_wrap_xml}
<w:jc w:val="left"/>
</w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="22"/></w:rPr><w:t>{text}</w:t></w:r>
</w:p>'''


def write_docx(path, body_xml):
    tmp = tempfile.mkdtemp(prefix='jj_')
    try:
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
            ("word/document.xml", body_xml),
        ]
        for relpath, content in files:
            full = os.path.join(tmp, relpath.replace("/", os.sep))
            os.makedirs(os.path.dirname(full), exist_ok=True)
            with open(full, "w", encoding="utf-8") as f:
                f.write(content)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, names in os.walk(tmp):
                for fn in names:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace(os.sep, '/')
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def render_pdf(word, docx, pdf):
    last = None
    for attempt in range(5):
        try:
            d = word.Documents.Open(docx, ReadOnly=True)
            time.sleep(0.4)
            d.SaveAs2(pdf, FileFormat=17)
            d.Close(SaveChanges=False)
            return True
        except Exception as e:
            last = e
            time.sleep(1.0 + attempt * 0.5)
    print(f"  ERR: {last}")
    return False


def measure_text_lines(pdf):
    """Get all text lines with their y position to detect line breaks."""
    d = fitz.open(pdf)
    page = d[0]
    blocks = page.get_text("blocks")
    lines = []
    for b in blocks:
        if b[6] == 0:  # text block
            txt = b[4].strip()
            if txt:
                lines.append({'x': b[0], 'y': b[1], 'text': txt})
    d.close()
    return lines


# Page width = 595.30pt, margins 42.55pt each side → content = 510pt
# But we'll use narrow-width tests by indenting
# For sz=22 (=11pt), Latin word "supercalifragilistic" ~= 100pt wide
# CJK string "あいうえおかきくけこさしすせそ" 15 chars = 165pt at 11pt fullwidth

VARIANTS = [
    # Test 1: Latin long word, narrow width (force wrap decision)
    ("V_JJ0_latin_default_wordWrap_on", para(text="supercalifragilisticexpialidocious", ind_xml='<w:ind w:right="6500"/>')),
    ("V_JJ1_latin_wordWrap_off", para(text="supercalifragilisticexpialidocious", ind_xml='<w:ind w:right="6500"/>', word_wrap_xml='<w:wordWrap w:val="off"/>')),
    # Test 2: CJK long string, narrow width
    ("V_JJ2_cjk_default_wordWrap_on", para(text="あいうえおかきくけこさしすせそたちつてと", ind_xml='<w:ind w:right="6500"/>')),
    ("V_JJ3_cjk_wordWrap_off", para(text="あいうえおかきくけこさしすせそたちつてと", ind_xml='<w:ind w:right="6500"/>', word_wrap_xml='<w:wordWrap w:val="off"/>')),
    # Test 3: Mixed
    ("V_JJ4_mixed_wordWrap_off", para(text="日本語ABCDEFGHIJKLMNOPQRSTUVWXYZ123あいうえお", ind_xml='<w:ind w:right="6500"/>', word_wrap_xml='<w:wordWrap w:val="off"/>')),
]


def main():
    pythoncom.CoInitialize()
    word = None
    for attempt in range(5):
        try:
            word = wc.Dispatch("Word.Application")
            time.sleep(2.0)
            word.Visible = False
            word.DisplayAlerts = False
            break
        except Exception as e:
            print(f"Word startup {attempt+1}: {e}")
            time.sleep(8.0)
    if word is None:
        print("Failed Word"); return
    print("V_JJ: wordWrap behavior on Latin vs CJK\n")
    print("Hypothesis (per ECMA): wordWrap=off → Latin no mid-word break, CJK still breaks")
    print()
    results = []
    try:
        for vid, paragraph in VARIANTS:
            print(f"=== {vid} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            write_docx(docx, doc_xml(paragraphs=paragraph))
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render"})
                continue
            lines = measure_text_lines(pdf)
            print(f"  {len(lines)} text blocks rendered:")
            for l in lines[:5]:
                print(f"    x={l['x']:.1f} y={l['y']:.1f} text={l['text'][:40]!r}")
            results.append({"id": vid, "lines": lines})
    finally:
        try: word.Quit()
        except: pass
    out = os.path.join(OUT_DIR, "results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
