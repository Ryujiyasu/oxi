# -*- coding: utf-8 -*-
"""Word's EMPTY-paragraph ¶-mark font resolution — controlled sweep.

mysignaiguide (no-grid, docDefaults ascii=hAnsi=eastAsia=cs=Arial, lang ja,
CJK body) measures its empty paragraphs at 1.5x fs (= a 1.297 CJK-ish natural
x the 1.15 line multiple), while Oxi's S583/S707 "prefer the ASCII font" path
returns Arial (1.1499) => every empty paragraph is ~2.5pt short.

The doc has NO ascii-vs-eastAsia conflict (both are Arial), so S583's
discriminator is not even engaged.  The open question is what Word does
AFTER the chain yields a LATIN-ONLY font in a CJK context: does it font-link
the mark to a CJK face (the S634 rule, but for the mark)?

Readout is SELF-CALIBRATING and immune to the 0.75pt COM quantization:
each case page holds

    P1 anchor "AAA"   (10pt Arial, explicit)
    P2 anchor "BBB"   <- gap(P1,P2) = anchor line height        [control]
    P3 EMPTY          (mark rPr = the swept variable, 20pt)
    P4 anchor "CCC"   <- gap(P2,P4) = anchor height + EMPTY height

    EMPTY_height = gap(P2,P4) - gap(P1,P2)

All spacing is zeroed and lineRule=auto line=240 (single), so the measured
advance IS the mark font's natural line height.  20pt amplifies the
discrimination: Arial 23.0 / Cambria 23.4 / MS Mincho 25.9 / Arial Unicode
MS ~26.0.

Usage: python _pb_markfont_gen.py gen | measure | read
"""
import os, sys, glob, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUTDIR = os.path.join(REPO, "pipeline_data", "_pb_markfont")

CJK_SENT = "本規程の適用範囲は、次の各号に定めるとおりとする。"

# (id, dd_ascii, dd_eastasia, lang_val, lang_ea, has_cjk_body, mark_rfonts)
CASES = [
    # 1. the mysignaiguide replica: everything Arial, ja, CJK body
    ("c1_arial_ja_cjk",     "Arial",     "Arial",      "ja",    None,    True,  ""),
    # 2. language discriminator (Latin lang)
    ("c2_arial_en_cjk",     "Arial",     "Arial",      "en-US", "en-US", True,  ""),
    # 3. document-CJK-content discriminator (no CJK anywhere)
    ("c3_arial_ja_nocjk",   "Arial",     "Arial",      "ja",    None,    False, ""),
    # 4. eastAsia is a REAL CJK font: ascii(Arial) or eastAsia(Mincho)?
    ("c4_ari_eaMin_ja",     "Arial",     "ＭＳ 明朝",  "ja",    None,    True,  ""),
    # 5. the S583 "model" shape: Latin ascii + CJK eastAsia (ascii won there)
    ("c5_camb_eaMin_ja",    "Cambria",   "ＭＳ 明朝",  "ja",    None,    True,  ""),
    # 6. reverse: CJK ascii + Latin eastAsia
    ("c6_minAscii_eaAri_ja","ＭＳ 明朝", "Arial",      "ja",    None,    True,  ""),
    # 7. hint=eastAsia on the mark
    ("c7_hint_ea",          "Arial",     "Arial",      "ja",    None,    True,
     '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial" w:hint="eastAsia"/>'),
    # 8. control: mark explicitly Arial Unicode MS (a CJK-capable Unicode face)
    ("c8_mark_aum",         "Arial",     "Arial",      "ja",    None,    True,
     '<w:rFonts w:ascii="Arial Unicode MS" w:hAnsi="Arial Unicode MS"'
     ' w:eastAsia="Arial Unicode MS" w:cs="Arial Unicode MS"/>'),
    # 9. control: mark explicitly MS Mincho
    ("c9_mark_mincho",      "Arial",     "Arial",      "ja",    None,    True,
     '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
     ' w:eastAsia="ＭＳ 明朝" w:cs="ＭＳ 明朝"/>'),
]

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOCRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat><w:compatSetting w:name="compatibilityMode"
 w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>"""


def styles_xml(ascii_f, ea_f, lang_val, lang_ea):
    lang = f'<w:lang w:val="{lang_val}"'
    if lang_ea:
        lang += f' w:eastAsia="{lang_ea}"'
    lang += "/>"
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="{ascii_f}" w:hAnsi="{ascii_f}" w:eastAsia="{ea_f}" w:cs="{ascii_f}"/>
<w:sz w:val="20"/><w:szCs w:val="20"/>{lang}
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr>
<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>"""


ARIAL_RPR = ('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>'
             '<w:sz w:val="20"/><w:szCs w:val="20"/>')


def anchor(text):
    return (f'<w:p><w:pPr><w:rPr>{ARIAL_RPR}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{ARIAL_RPR}</w:rPr><w:t>{text}</w:t></w:r></w:p>')


def cjk_para():
    return (f'<w:p><w:r><w:t>{CJK_SENT}</w:t></w:r></w:p>')


def empty_para(mark_rfonts):
    # the mark's own rPr: 20pt, plus the swept rFonts (may be empty = inherit)
    return (f'<w:p><w:pPr><w:rPr>{mark_rfonts}'
            f'<w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:pPr></w:p>')


SECTPR = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
          '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
          ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for (cid, a, ea, lv, le, cjk, mark) in CASES:
        body = []
        if cjk:
            body.append(cjk_para())
        body.append(anchor("AAA"))
        body.append(anchor("BBB"))
        body.append(empty_para(mark))
        body.append(anchor("CCC"))
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
               '<w:body>' + "".join(body) + SECTPR + '</w:body></w:document>')
        path = os.path.join(OUTDIR, cid + ".docx")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", DOCRELS)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", styles_xml(a, ea, lv, le))
            z.writestr("word/settings.xml", SETTINGS)
        print("gen", cid)


def measure():
    import win32com.client as win32
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        for path in sorted(glob.glob(os.path.join(OUTDIR, "*.docx"))):
            pdf = path[:-5] + ".pdf"
            d = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
            try:
                d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf),
                                      ExportFormat=17)
                print("measured", os.path.basename(path))
            finally:
                d.Close(False)
    finally:
        word.Quit()


def read():
    import fitz
    print(f"{'case':<24} {'anchor_h':>9} {'empty_h':>9}   verdict")
    for path in sorted(glob.glob(os.path.join(OUTDIR, "*.pdf"))):
        cid = os.path.basename(path)[:-4]
        doc = fitz.open(path)
        bl = {}
        for pg in doc:
            for b in pg.get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    t = "".join(s["text"] for s in l["spans"]).strip()
                    if t in ("AAA", "BBB", "CCC"):
                        bl[t] = l["spans"][0]["origin"][1]
        if len(bl) < 3:
            print(f"{cid:<24}  MISSING anchors {sorted(bl)}")
            continue
        anchor_h = bl["BBB"] - bl["AAA"]
        empty_h = (bl["CCC"] - bl["BBB"]) - anchor_h
        # 20pt candidates
        cands = {"Arial 1.1499": 22.998, "Cambria 1.1724": 23.448,
                 "Calibri 1.2207": 24.414, "MSMincho 1.297": 25.94,
                 "AUM ~1.298": 25.96}
        best = min(cands.items(), key=lambda kv: abs(kv[1] - empty_h))
        print(f"{cid:<24} {anchor_h:9.3f} {empty_h:9.3f}   "
              f"{best[0]} (d={empty_h-best[1]:+.2f})")




# ---- round 2: theme-Jpan fallback + lang override --------------------------

# A stubbed theme (empty clrScheme/fmtScheme) is schema-invalid and Word
# refuses to open the file, so round 2 transplants a REAL theme part and only
# swaps its <a:font script="Jpan"> face — the faithful-host lesson (S931).
THEME_SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                         "mysignaiguide.docx")

CASES2 = [
    # (id, jpan_theme_face, mark_extra_rpr_AFTER_sz, line)
    ("t1_theme_mincho",  "ＭＳ 明朝", "", "240"),
    ("t2_theme_gothic",  "ＭＳ ゴシック", "", "240"),
    ("t5_theme_meiryo",  "Meiryo",                "", "240"),
    ("t4_theme_mult115", "ＭＳ 明朝", "", "276"),
    # the mark's OWN language (correct element order: after sz/szCs)
    ("u2_marklang_en",   "ＭＳ 明朝", '<w:lang w:val="en-US"/>', "240"),
]

# docDefaults-language variants (the decisive discriminator), theme present
CASES3 = [
    # (id, dd_lang_val, dd_lang_eastAsia)
    ("u1_theme_langEN",   "en-US", "en-US"),
    ("u4_theme_normalJP", "en-US", "ja-JP"),   # the NORMAL Japanese config
]


def _theme_with(jpan):
    import re
    t = zipfile.ZipFile(THEME_SRC).read("word/theme/theme1.xml").decode("utf8")
    return re.sub(r'(<a:font script="Jpan" typeface=")[^"]*(")',
                  lambda m: m.group(1) + jpan + m.group(2), t)


def _write_themed(cid, styles, mark_extra, jpan):
    body = [cjk_para(), anchor("AAA"), anchor("BBB"),
            f'<w:p><w:pPr><w:rPr><w:sz w:val="40"/><w:szCs w:val="40"/>'
            f'{mark_extra}</w:rPr></w:pPr></w:p>',
            anchor("CCC")]
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>' + "".join(body) + SECTPR + '</w:body></w:document>')
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/theme/theme1.xml"'
                    ' ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
                    "</Types>")
    dr = DOCRELS.replace("</Relationships>",
                         '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/'
                         'officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
                         "</Relationships>")
    with zipfile.ZipFile(os.path.join(OUTDIR, cid + ".docx"), "w",
                         zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", dr)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/theme/theme1.xml", _theme_with(jpan))
    print("gen2", cid)


def gen2():
    os.makedirs(OUTDIR, exist_ok=True)
    for (cid, jpan, extra, line) in CASES2:
        st = styles_xml("Arial", "Arial", "ja", None).replace(
            'w:line="240"', f'w:line="{line}"')
        _write_themed(cid, st, extra, jpan)
    for (cid, lv, le) in CASES3:
        _write_themed(cid, styles_xml("Arial", "Arial", lv, le), "",
                      "ＭＳ 明朝")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    {"gen": gen, "gen2": gen2, "measure": measure, "read": read}[cmd]()
