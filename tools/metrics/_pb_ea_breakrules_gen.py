# -*- coding: utf-8 -*-
"""What does <w:doNotUseEastAsianBreakRules/> do in the 2026-10 Word?

ja/technical__c6d7cdb1 (compat 14, settings carry doNotUseEastAsianBreakRules,
Body Text style lang eastAsia=en-US, 5pt MS Mincho): Word ends a line at
«…表示されます» and starts the next line with «。» alone; Oxi keeps «。» on the
line (kinsoku burasage) and the page holds one line more.

Arms: setting on/off x run lang eastAsia ja-JP / en-US x compat 14/15, a Japanese
sentence ending in «。» followed by more text, right indent swept so «。» meets
the line end.  Per arm: does a line START with «。» (no kinsoku), does «。» END a
line past the right edge (hang), or neither (the previous character is pushed).

    python tools/metrics/_pb_ea_breakrules_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_ea_breakrules")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
TEXT = "スリープタイマーが設定されると、ディスプレイに表示されます。次の行の文章です"
TEXT = os.environ.get("PB_TEXT", TEXT)
MARK = os.environ.get("PB_MARK", "。")   # the punctuation under test   # e.g. the real line with spaces and Latin «LCD»
RIGHTS = list(range(0, 2400, 10))
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')


def build(name, rule_off, lang, cm, ea="ＭＳ 明朝"):
    paras = []
    for i, r in enumerate(RIGHTS):
        pb = "<w:pageBreakBefore/>" if i else ""
        paras.append('<w:p><w:pPr>%s<w:ind w:right="%d"/><w:spacing w:before="0" w:after="0"/></w:pPr>'
                     '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="%s"/><w:sz w:val="21"/>'
                     '<w:lang w:eastAsia="%s"/></w:rPr><w:t>%s</w:t></w:r></w:p>' % (pb, r, ea, lang, TEXT))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="6000" w:h="4000"/><w:pgMar w:top="400" w:bottom="400" w:left="400" w:right="400"/></w:sectPr>'
           '</w:body></w:document>' % (W, "".join(paras)))
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:compat>%s'
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/>'
                '</w:compat></w:settings>' % (W, "<w:doNotUseEastAsianBreakRules/>" if rule_off else "", cm))
    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, name + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/document.xml", doc); z.writestr("word/settings.xml", settings)
    return p


def word_classes(docx):
    import fitz, win32com.client
    pdf = docx[:-5] + ".pdf"
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    edge_page = 300 - 20
    seen = {"starts_line": 0, "hangs": 0, "pushed": 0}
    for r, pg in zip(RIGHTS, fitz.open(pdf)):
        lim = edge_page - r / 20
        lines = []
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                ch = [c for s in l["spans"] for c in s["chars"] if c["c"].strip()]
                if ch:
                    lines.append(ch)
        for k, ch in enumerate(lines):
            # only the mark after «表示されます»: the previous line ends with す
            if ch[0]["c"] == MARK and k > 0 and lines[k - 1][-1]["c"] == "す":
                seen["starts_line"] += 1
            if ch[-1]["c"] == MARK and len(ch) > 1 and ch[-2]["c"] == "す" and ch[-1]["bbox"][2] > lim + 0.5:
                seen["hangs"] += 1
            if k + 1 < len(lines) and lines[k + 1][0]["c"] != MARK and "".join(c["c"] for c in lines[k + 1]).startswith("す" + MARK):
                seen["pushed"] += 1
    return seen


if __name__ == "__main__":
    # PB_EAFONTS: the run's w:eastAsia faces to try (a Latin face here = the CJK text falls back)
    for ea in os.environ.get("PB_EAFONTS", "ＭＳ 明朝").split(","):
        for rule_off in ((False, True) if not os.environ.get("PB_RULE_ON_ONLY") else (False,)):
            for lang in os.environ.get("PB_LANGS", "ja-JP,en-US").split(","):
                cm = 14
                name = "rule%s_%s_cm%d_%s" % ("OFF" if rule_off else "on", lang, cm, ea.replace(" ", ""))
                print("%-34s %s" % (name, word_classes(build(name, rule_off, lang, cm, ea))), flush=True)
