# -*- coding: utf-8 -*-
"""V_HH: Measure Word's tab positioning in hanging-indent paragraphs."""
import os, sys, time, json, zipfile, shutil, tempfile
import pythoncom, win32com.client as wc
import fitz

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES, RELS_ROOT, WORD_RELS

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/v_hh_tab_hanging")
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
</w:settings>'''


def doc_xml(*, ind_xml="", marker="エ", body="ABC", explicit_tabs_xml=""):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>
<w:p>
<w:pPr>
{ind_xml}
{explicit_tabs_xml}
<w:jc w:val="left"/>
</w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="22"/></w:rPr><w:t>{marker}</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="22"/></w:rPr><w:tab/></w:r>
<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="22"/></w:rPr><w:t>{body}</w:t></w:r>
</w:p>
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838" w:code="9"/>
<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>'''


def write_docx(path, **kwargs):
    tmp = tempfile.mkdtemp(prefix='hh_')
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


def measure(pdf, chars):
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


# Page left margin = 851 twips = 42.55pt
# defaultTabStop = 840 twips = 42pt

VARIANTS = [
    # (id, marker, body, ind_xml, explicit_tabs_xml, description)
    ("V_HH0_no_ind", "エ", "本文", "", "",
     "no ind, no tabs def: tab from after marker → next 42pt multiple"),
    ("V_HH1_hang_140_left_564", "エ", "本文",
     '<w:ind w:left="564" w:hanging="140"/>', "",
     "ind left=564 hanging=140: marker at 21.2pt, tab → ?"),
    ("V_HH2_hang_280_left_564", "エ", "本文",
     '<w:ind w:left="564" w:hanging="280"/>', "",
     "ind left=564 hanging=280: marker at 14pt, tab → 28.2pt or 42pt?"),
    ("V_HH3_explicit_tab_at_50pt", "エ", "本文",
     "", '<w:tabs><w:tab w:val="left" w:pos="1000"/></w:tabs>',
     "explicit tab stop at 1000 twips = 50pt"),
    ("V_HH4_hang_AND_explicit_tab", "エ", "本文",
     '<w:ind w:left="564" w:hanging="140"/>',
     '<w:tabs><w:tab w:val="left" w:pos="1000"/></w:tabs>',
     "hanging + explicit tab at 50pt"),
    ("V_HH5_left_only_no_hanging", "エ", "本文",
     '<w:ind w:left="564"/>', "",
     "ind left=564 only (no hanging): marker at 28.2pt, tab → 42pt or 84pt?"),
    ("V_HH6_marker_two_chars", "ＡＢ", "本文",
     '<w:ind w:left="564" w:hanging="140"/>', "",
     "marker is 2 fullwidth chars: marker takes 22pt, tab → ?"),
    ("V_HH7_no_hanging_explicit_tab", "エ", "本文",
     "", '<w:tabs><w:tab w:val="left" w:pos="282"/></w:tabs>',
     "explicit tab at 14.1pt (BEFORE marker end at ~22pt)"),
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
    print("V_HH: tab positioning in various paragraph configs")
    print(f"page left margin = 42.55pt; defaultTabStop = 42pt")
    print(f"sz=22 = 11pt, marker 'エ' fullwidth = 11pt wide\n")
    results = []
    try:
        for vid, marker, body, ind_xml, tabs_xml, desc in VARIANTS:
            print(f"=== {vid}: {desc} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            write_docx(docx, marker=marker, body=body, ind_xml=ind_xml, explicit_tabs_xml=tabs_xml)
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render"})
                continue
            chars = list(set(marker + body[:1]))
            positions = measure(pdf, chars)
            for ch in chars:
                if ch in positions and positions[ch]:
                    p = positions[ch][0]
                    print(f"  '{ch}' at x=[{p['x']:.2f}, {p['x1']:.2f}]  y=[{p['y']:.2f}, {p['y1']:.2f}]")
            results.append({"id": vid, "marker": marker, "body": body,
                          "ind_xml": ind_xml, "explicit_tabs_xml": tabs_xml,
                          "positions": positions})
    finally:
        try: word.Quit()
        except: pass
    out = os.path.join(OUT_DIR, "results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
