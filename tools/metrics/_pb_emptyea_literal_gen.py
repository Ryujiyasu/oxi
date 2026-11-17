# -*- coding: utf-8 -*-
"""A run that says eastAsiaTheme="minorEastAsia" against a theme whose minor
<a:ea> is EMPTY, while a STYLE above it names a literal East Asian face.

S1397 reads the empty slot as "no font" and lets the style's literal through.
Its evidence (legal__0676305d) had themeFontLang zh-TW, whose Hant face 新細明體
is also the literal Normal names, so it cannot tell the two readings apart.
legal__02f84965 can: Normal ＭＳ Ｐ明朝, theme Jpan ＭＳ 明朝, Word paints MS-Mincho.

    python _pb_emptyea_literal_gen.py gen  OUTDIR
    python _pb_emptyea_literal_gen.py pdf  OUTDIR    # Word truth (PDF span faces)
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402
from _pb_themeea_gen import theme_xml  # noqa: E402

JA = 'w:val="en-US" w:eastAsia="ja-JP"'
TW = 'w:val="en-US" w:eastAsia="zh-TW"'
REF = '<w:rFonts w:asciiTheme="minorEastAsia" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorEastAsia"/>'
DEFREF = '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi"/>'


def lit(face):
    return '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>' % (face, face, face)


# (label, docDefaults rFonts, Normal rFonts, run rFonts, theme scripts, themeFontLang, <a:ea>)
JPAN = {"Jpan": "ＭＳ 明朝"}
ARMS = [
    ("normal_pmincho_runref", DEFREF, lit("ＭＳ Ｐ明朝"), REF, JPAN, JA, ""),
    ("normal_pmincho_noref", DEFREF, lit("ＭＳ Ｐ明朝"), "", JPAN, JA, ""),
    ("normal_gothic_runref", DEFREF, lit("ＭＳ ゴシック"), REF, JPAN, JA, ""),
    ("normal_pmincho_runref_yu", DEFREF, lit("ＭＳ Ｐ明朝"), REF, {"Jpan": "游明朝"}, JA, ""),
    ("docdef_gothic_runref", lit("ＭＳ ゴシック"), "", REF, JPAN, JA, ""),
    ("docdef_gothic_normal_pm_runref", lit("ＭＳ ゴシック"), lit("ＭＳ Ｐ明朝"), REF, JPAN, JA, ""),
    ("normal_pmincho_runref_named_ea", DEFREF, lit("ＭＳ Ｐ明朝"), REF, JPAN, JA, "ＭＳ ゴシック"),
    ("tw_normal_pmincho_runref", DEFREF, lit("ＭＳ Ｐ明朝"), REF,
     {"Jpan": "ＭＳ 明朝", "Hant": "新細明體"}, TW, ""),
    ("normal_pmincho_runref_nojpan", DEFREF, lit("ＭＳ Ｐ明朝"), REF, {}, JA, ""),
    # the reference on the paragraph STYLE, a literal below it in docDefaults
    ("docdef_gothic_normalref_norun", lit("ＭＳ ゴシック"), REF, "", JPAN, JA, ""),
    # themeFontLang zh-TW with no Hant entry in the theme
    ("tw_nohant_normal_pmincho_runref", DEFREF, lit("ＭＳ Ｐ明朝"), REF, JPAN, TW, ""),
]
TEXT = "（様式Ｃ－２）本文の東亜フォントはどれか「質問」、ＯＳ"


def docx(out, label):
    return os.path.join(out, "emptyea_%s.docx" % label)


def gen(out):
    os.makedirs(out, exist_ok=True)
    ct = CT.replace(
        "</Types>",
        '<Override PartName="/word/theme/theme1.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    rel = ('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
           'officeDocument/2006/relationships/%s" Target="%s"/>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             + rel % (1, "styles", "styles.xml") + rel % (2, "theme", "theme/theme1.xml")
             + rel % (3, "settings", "settings.xml") + "</Relationships>")
    for label, dflt, normal, run, scripts, lang, ea in ARMS:
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  "<w:docDefaults><w:rPrDefault><w:rPr>" + dflt + '<w:sz w:val="21"/>'
                  '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
                  '<w:name w:val="Normal"/><w:rPr>' + normal + "</w:rPr></w:style></w:styles>")
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                    '<w:compat><w:compatSetting w:name="compatibilityMode"'
                    ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                    "<w:themeFontLang %s/></w:settings>" % lang)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body><w:p><w:r><w:rPr>" + run + '<w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>'
               + TEXT + "</w:t></w:r></w:p>"
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
               '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
               "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(out, label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/theme/theme1.xml", theme_xml(ea, scripts))
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), out))


def pdf(out):
    import fitz
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for arm in ARMS:
            label = arm[0]
            src = docx(out, label)
            dst = src.replace(".docx", ".pdf")
            d = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
            try:
                d.ExportAsFixedFormat(dst, 17)
                com = str(d.Paragraphs(1).Range.Font.NameFarEast)
            finally:
                d.Close(False)
            faces = sorted({e[3].split("+")[-1] for e in fitz.open(dst)[0].get_fonts(False)})
            print("%-34s %-30s COM=%r" % (label, ", ".join(faces), com))
    finally:
        app.Quit()


if __name__ == "__main__":
    cmd, out = sys.argv[1], sys.argv[2]
    {"gen": gen, "pdf": pdf}[cmd](out)
