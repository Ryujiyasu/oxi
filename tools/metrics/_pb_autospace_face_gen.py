# -*- coding: utf-8 -*-
"""Does Word put the East Asian / Latin gap (autoSpaceDE) around Latin letters whose
ASCII face is a Japanese face?

golden mysignaiguide: every run is Arial Unicode MS (absent here), docDefaults lang ja;
Word draws «mysign» / «AI» in MS-Mincho half-width and puts NO gap at «、m» «n公»
«成A» «Iを» (advances 9.98 / 4.95 / 9.96 / 5.04 at 10pt).  Oxi adds 1.25-2.5.

Arms: run rFonts (ascii=hAnsi=eastAsia) x fontTable declaration x run language.
Each doc is opened in its own new Word process (DispatchEx).  Prints, per arm, the
advance of each boundary pair read from the PDF.

    python tools/metrics/_pb_autospace_face_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_autospace_face")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
# PB_COMPAT=15 adds a settings part with that compatibilityMode (none = no settings part)
COMPAT = os.environ.get("PB_COMPAT")
# PB_GRID=1 adds <w:docGrid w:type="lines" w:linePitch="360"/>; PB_CSC=1 adds
# <w:characterSpacingControl w:val="compressPunctuation"/> (settings part forced)
GRID = os.environ.get("PB_GRID")
CSC = os.environ.get("PB_CSC")
# PB_DD_EA=<face>: docDefaults rFonts eastAsia (default Arial, as mysignaiguide)
DD_EA = os.environ.get("PB_DD_EA", "Arial")
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>'
         '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
TEXT = "本ガイドラインは、mysign公式サイト生成AIを利用SEO効果"
# arm -> (face, fontTable charset or None, docDefaults lang)
ARMS = {
    "mincho_ja": ("ＭＳ 明朝", None, "ja"),
    "mincho_en": ("ＭＳ 明朝", None, "en-US"),
    "gothic_ja": ("ＭＳ ゴシック", None, "ja"),
    "aum_ja": ("Arial Unicode MS", None, "ja"),
    "aum80_ja": ("Arial Unicode MS", "80", "ja"),
    "aum_en": ("Arial Unicode MS", None, "en-US"),
    "arial_ja": ("Arial", None, "ja"),
    "century_ja": ("Century", None, "ja"),
    "yumin_ja": ("游明朝", None, "ja"),
    # lang = (w:val, w:eastAsia)
    "mincho_jaJP": ("ＭＳ 明朝", None, ("ja", "ja-JP")),
    "arial_enJP": ("Arial", None, ("en-US", "ja-JP")),
    "aum_jaJP": ("Arial Unicode MS", None, ("ja", "ja-JP")),
    "aum_enJP": ("Arial Unicode MS", None, ("en-US", "ja-JP")),
    "mincho_enJP": ("ＭＳ 明朝", None, ("en-US", "ja-JP")),
    "arial_enUS": ("Arial", None, ("en-US", "en-US")),
    # no <w:lang> in docDefaults (Normal.dotm's language)
    "mincho_nolang": ("ＭＳ 明朝", None, None),
    "arial_nolang": ("Arial", None, None),
    "aum_nolang": ("Arial Unicode MS", None, None),
    "mincho_jaJPjaJP": ("ＭＳ 明朝", None, ("ja-JP", "ja-JP")),
    "arial_jaJP": ("Arial", None, ("ja", "ja-JP")),
    "century_enJP": ("Century", None, ("en-US", "ja-JP")),
    "yumin_enJP": ("游明朝", None, ("en-US", "ja-JP")),
    "mincho_zhCN": ("ＭＳ 明朝", None, ("zh-CN", "ja-JP")),
    "mincho_frJP": ("ＭＳ 明朝", None, ("fr-FR", "ja-JP")),
    # face = (ascii/hAnsi, eastAsia)
    "arialYG_ja": (("Arial", "游ゴシック"), None, ("ja", "ja-JP")),
    "arialYG_en": (("Arial", "游ゴシック"), None, ("en-US", "ja-JP")),
    "arialMin_ko": (("Arial", "ＭＳ 明朝"), None, ("ko-KR", "ja-JP")),
    "arialMin_zhTW": (("Arial", "ＭＳ 明朝"), None, ("zh-TW", "ja-JP")),
    "arialMin_ja": (("Arial", "ＭＳ 明朝"), None, ("ja", "ja-JP")),
    # the language on the RUN (docDefaults en-US / ja-JP)
    "arialMin_jaRun": (("Arial", "ＭＳ 明朝"), None, ("en-US", "ja-JP"), "ja-JP"),
    "arialMin_jaRunShort": (("Arial", "ＭＳ 明朝"), None, ("en-US", "ja-JP"), "ja"),
    "arialMin_jaJPdd": (("Arial", "ＭＳ 明朝"), None, ("ja-JP", "ja-JP")),
    "arialMin_jaNoEA": (("Arial", "ＭＳ 明朝"), None, "ja"),
}


def build(arm):
    face, charset, lang = ARMS[arm][:3]
    run_lang = ARMS[arm][3] if len(ARMS[arm]) > 3 else None
    asc, ea = face if isinstance(face, tuple) else (face, face)
    langxml = "" if lang is None else ('<w:lang w:val="%s"/>' % lang) if isinstance(lang, str) else ('<w:lang w:val="%s" w:eastAsia="%s"/>' % lang)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"%s w:cs="Arial"/><w:sz w:val="20"/>%s'
              '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') % (W, "" if DD_EA == "none" else ' w:eastAsia="%s"' % DD_EA, langxml)
    cs = '<w:charset w:val="%s"/>' % charset if charset else ""
    fonttable = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts %s>'
                 '<w:font w:name="Arial Unicode MS">%s<w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font></w:fonts>') % (W, cs)
    rpr = '<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/><w:sz w:val="20"/>%s</w:rPr>' % (
        asc, asc, ea, asc, '<w:lang w:val="%s"/>' % run_lang if run_lang else "")
    body = '<w:p><w:r>%s<w:t>%s</w:t></w:r></w:p>' % (rpr, TEXT)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>%s</w:sectPr></w:body></w:document>'
           % (W, body, '<w:docGrid w:type="lines" w:linePitch="360"/>' if GRID else ""))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "as_%s%s%s%s%s.docx" % (arm, "_c" + COMPAT if COMPAT else "", "_g" if GRID else "", "_csc" if CSC else "", "" if DD_EA == "Arial" else "_ddea"))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", styles)
        z.writestr("word/fontTable.xml", fonttable); z.writestr("word/document.xml", doc)
        if COMPAT or CSC:
            z.writestr("word/settings.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s>%s<w:compat>%s'
                       '</w:compat></w:settings>' % (W,
                       '<w:characterSpacingControl w:val="compressPunctuation"/>' if CSC else "",
                       '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%s"/>' % COMPAT if COMPAT else ""))
    return path


def pairs(pdf):
    import pymupdf
    chars = []
    for b in pymupdf.open(pdf)[0].get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    chars.append((c["c"], c["origin"][0], s["font"]))
    out = {}
    for a, b in zip(chars, chars[1:]):
        if a[0] + b[0] in ("、m", "mn"[:0] + "n公", "成A", "Iを", "用S", "O効", "my"):
            out[a[0] + b[0]] = round(b[1] - a[1], 2)
    return out, sorted({c[2] for c in chars})


if __name__ == "__main__":
    import win32com.client
    for arm in [a for a in ARMS if len(sys.argv) < 2 or a in sys.argv[1:]]:
        p = build(arm); pdf = p[:-5] + ".pdf"
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
        adv, faces = pairs(pdf)
        print("%-11s %s %s" % (arm, adv, faces), flush=True)
