# -*- coding: utf-8 -*-
"""V_GG: Minimal repro of vertical cell (tbRlV) — measure Word's rendering rules.

Build a single-row table with one vertical cell:
- textDirection w:val="tbRlV"
- Various text contents (CJK, Latin, mixed, punctuation)
- Various cell widths and heights
- vAlign options

Measure each character's pixel position via PDF + fitz, derive layout rules.
"""
import os, sys, time, json, zipfile, shutil, tempfile
import pythoncom, win32com.client as wc
import fitz

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES, RELS_ROOT, WORD_RELS

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/v_gg_vertical")
os.makedirs(OUT_DIR, exist_ok=True)

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>
<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:defaultTabStop w:val="840"/>
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
</w:compat>
</w:styles>'''.replace('</w:styles>', '</w:settings>')


def doc_xml(*, vert_text, cell_w_dxa=600, row_h_dxa=2000, valign="center", text_dir="tbRlV"):
    valign_xml = f'<w:vAlign w:val="{valign}"/>' if valign else ''
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>
<w:tbl>
<w:tblPr>
<w:tblW w:w="0" w:type="auto"/>
<w:tblLayout w:type="fixed"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:color="auto"/>
<w:insideV w:val="single" w:sz="4" w:color="auto"/>
</w:tblBorders>
</w:tblPr>
<w:tblGrid>
<w:gridCol w:w="{cell_w_dxa}"/>
<w:gridCol w:w="3000"/>
</w:tblGrid>
<w:tr>
<w:trPr><w:trHeight w:val="{row_h_dxa}" w:hRule="atLeast"/></w:trPr>
<w:tc>
<w:tcPr>
<w:tcW w:w="{cell_w_dxa}" w:type="dxa"/>
<w:textDirection w:val="{text_dir}"/>
{valign_xml}
</w:tcPr>
<w:p>
<w:pPr><w:jc w:val="center"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>{vert_text}</w:t></w:r>
</w:p>
</w:tc>
<w:tc>
<w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
<w:p><w:r><w:t>右セル: {vert_text}</w:t></w:r></w:p>
</w:tc>
</w:tr>
</w:tbl>
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838" w:code="9"/>
<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>'''


def write_docx(path, **kwargs):
    tmp = tempfile.mkdtemp(prefix='gg_')
    try:
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
            ("word/document.xml", doc_xml(**kwargs)),
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


def measure_chars(pdf, chars):
    """Find positions of each char in chars list."""
    d = fitz.open(pdf)
    page = d[0]
    res = {}
    for ch in chars:
        positions = []
        for inst in page.search_for(ch):
            positions.append({"x": inst.x0, "y": inst.y0, "x1": inst.x1, "y1": inst.y1})
        res[ch] = positions
    d.close()
    return res


VARIANTS = [
    ("V_GG0_simple_3char", {"vert_text": "連絡先", "cell_w_dxa": 600}),
    ("V_GG1_5char", {"vert_text": "東京都港区", "cell_w_dxa": 600}),
    ("V_GG2_with_latin", {"vert_text": "ABC１２３", "cell_w_dxa": 700}),  # Latin + fullwidth digits
    ("V_GG3_with_punct", {"vert_text": "あ、い。う", "cell_w_dxa": 700}),
    ("V_GG4_wider_cell_900", {"vert_text": "連絡先", "cell_w_dxa": 900}),
    ("V_GG5_taller_row_3000", {"vert_text": "連絡先", "row_h_dxa": 3000}),
    ("V_GG6_valign_top", {"vert_text": "連絡先", "valign": "top"}),
    ("V_GG7_valign_bottom", {"vert_text": "連絡先", "valign": "bottom"}),
    ("V_GG8_btLr", {"vert_text": "連絡先", "text_dir": "btLr"}),  # bottom-to-top
    ("V_GG9_horizontal_control", {"vert_text": "連絡先", "text_dir": "lrTb"}),  # control: horizontal
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
    print("V_GG: vertical cell rendering measurement\n")
    results = []
    try:
        for vid, kwargs in VARIANTS:
            print(f"=== {vid} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            try:
                write_docx(docx, **kwargs)
            except Exception as e:
                print(f"  build failed: {e}")
                continue
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render"})
                continue
            chars = list(kwargs.get('vert_text', ''))
            positions = measure_chars(pdf, chars)
            print(f"  text: {kwargs.get('vert_text', '')!r}")
            for ch in chars:
                if ch in positions and positions[ch]:
                    p = positions[ch][0]  # first occurrence
                    print(f"    '{ch}' at x=[{p['x']:.2f}, {p['x1']:.2f}] y=[{p['y']:.2f}, {p['y1']:.2f}]")
            results.append({"id": vid, "kwargs": kwargs, "positions": positions})
    finally:
        try: word.Quit()
        except: pass
    out = os.path.join(OUT_DIR, "results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
