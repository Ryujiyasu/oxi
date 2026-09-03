# -*- coding: utf-8 -*-
"""What does an EMPTY `<a:ea typeface=""/>` mean for the theme's EA font?

S323 reads it as "the EA slot is explicitly empty -- do NOT fall through to the
`<a:font script="Jpan">` entry, use rPrDefault instead", and suppresses the Jpan
face for the MINOR (body) font. Its two evidence documents both inherit a
LITERAL `w:eastAsia` from docDefaults, so the theme is never consulted in either
of them (`_theme_ea_reach.py`: 0 CJK characters resolve through the theme in
both) -- they cannot answer the question they were used to answer. S492e/S845
already re-derived the MAJOR half the other way.

These arms vary ONE thing at a time: what the minor font's `<a:ea>` says, and
whether a Jpan entry sits beside it. The body text names NO font at all, so the
theme is the only possible source for its face.

    python _pb_themeea_gen.py gen
    python _pb_themeea_gen.py pdf      # Word truth: the face in the exported PDF
    python _pb_themeea_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_themeea")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

ANS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'

# (label, minor <a:ea> typeface, {script: face} supplemental entries,
#  themeFontLang attrs or None to omit, the RUN's w:lang w:eastAsia or None).
# ★An arm that OMITS <a:ea> entirely was tried and Word refuses to open the file
# ("this file may be corrupt") -- the element is required by the schema, so
# "absent" is not a state a real document can be in.
JA = 'w:val="en-US" w:eastAsia="ja-JP"'
JPAN = {"Jpan": "游明朝"}
ARMS = [
    ("empty_ea_jpan_mincho", "", JPAN, JA, None),
    ("empty_ea_jpan_gothic", "", {"Jpan": "游ゴシック"}, JA, None),
    ("empty_ea_no_jpan", "", {}, JA, None),
    ("named_ea_jpan_mincho", "メイリオ", JPAN, JA, None),
    ("named_ea_jpan_gothic", "メイリオ", {"Jpan": "游ゴシック"}, JA, None),
    ("named_ea_no_jpan", "メイリオ", {}, JA, None),
    ("empty_ea_jpan_nolang", "", JPAN, None, None),
    ("empty_ea_jpan_enlang", "", JPAN, 'w:val="en-US"', None),
    # ★The run's own East Asian LANGUAGE picks which script entry applies.
    # d1e8ac8's single theme-resolved run carries w:lang w:eastAsia="zh-TW"
    # beside a Jpan-only theme -- which is why Word answered ＭＳ 明朝 there and
    # why S323 read that as "an empty <a:ea> suppresses Jpan".
    ("runlang_ja_jpan", "", JPAN, JA, "ja-JP"),
    ("runlang_zhTW_jpan", "", JPAN, JA, "zh-TW"),
    ("runlang_zhCN_jpan", "", JPAN, JA, "zh-CN"),
    ("runlang_ko_jpan", "", JPAN, JA, "ko-KR"),
    ("runlang_zhTW_hant", "", {"Jpan": "游明朝", "Hant": "ＭＳ ゴシック"}, JA, "zh-TW"),
    ("runlang_zhTW_named_ea", "メイリオ", JPAN, JA, "zh-TW"),
    ("runlang_ja_themelang_zhTW", "", JPAN, 'w:val="en-US" w:eastAsia="zh-TW"', "ja-JP"),
    ("nolang_themelang_zhTW", "", JPAN, 'w:val="en-US" w:eastAsia="zh-TW"', None),
]
TEXT = "本文の東亜フォントはどれか"


def docx(label):
    return os.path.join(OUT, "themeea_%s.docx" % label)


def font_slot(kind, latin, ea_val, scripts):
    ea_el = "" if ea_val is None else '<a:ea typeface="%s"/>' % ea_val
    fonts = "".join('<a:font script="%s" typeface="%s"/>' % (s, f)
                    for s, f in sorted((scripts or {}).items()))
    return ('<a:%sFont><a:latin typeface="%s"/>%s<a:cs typeface=""/>%s</a:%sFont>'
            % (kind, latin, ea_el, fonts, kind))


def theme_xml(ea, scripts):
    clr = "".join(
        '<a:%s><a:srgbClr val="%s"/></a:%s>' % (n, v, n) for n, v in
        [("dk2", "44546A"), ("lt2", "E7E6E6"), ("accent1", "4472C4"),
         ("accent2", "ED7D31"), ("accent3", "A5A5A5"), ("accent4", "FFC000"),
         ("accent5", "5B9BD5"), ("accent6", "70AD47"), ("hlink", "0563C1"),
         ("folHlink", "954F72")])
    fill3 = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    ln3 = '<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' * 3
    eff3 = '<a:effectStyle><a:effectLst/></a:effectStyle>' * 3
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<a:theme %s name="probe"><a:themeElements>'
            '<a:clrScheme name="p">'
            '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
            '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>%s</a:clrScheme>'
            '<a:fontScheme name="p">%s%s</a:fontScheme>'
            '<a:fmtScheme name="p"><a:fillStyleLst>%s</a:fillStyleLst>'
            '<a:lnStyleLst>%s</a:lnStyleLst>'
            '<a:effectStyleLst>%s</a:effectStyleLst>'
            '<a:bgFillStyleLst>%s</a:bgFillStyleLst></a:fmtScheme>'
            '</a:themeElements></a:theme>'
            % (ANS, clr,
               font_slot("major", "Century", "", {"Jpan": "游ゴシック Light"}),
               font_slot("minor", "Century", ea, scripts),
               fill3, ln3, eff3, fill3))


def gen():
    os.makedirs(OUT, exist_ok=True)
    ct = CT.replace(
        "</Types>",
        '<Override PartName="/word/theme/theme1.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    rel = ('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
           'officeDocument/2006/relationships/%s" Target="%s"/>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
             'relationships">'
             + rel % (1, "styles", "styles.xml")
             + rel % (2, "theme", "theme/theme1.xml")
             + rel % (3, "settings", "settings.xml")
             + "</Relationships>")
    # docDefaults hands out the THEME token and nothing downstream names a face,
    # so the theme is the only possible source for the body's East Asian font.
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" '
              'w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="21"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/></w:style></w:styles>')
    # ★A probe that omits settings.xml gets a DIFFERENT resolution (Word drops to
    # a legacy compatibility mode), and themeFontLang is what names the script
    # whose supplemental <a:font> entry applies. Both must be present.
    def settings_for(lang):
        tfl = "" if lang is None else "<w:themeFontLang %s/>" % lang
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                '<w:compat><w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                + tfl + "</w:settings>")
    def doc_for(runlang):
        rpr = ("" if runlang is None else
               '<w:rPr><w:lang w:eastAsia="%s"/></w:rPr>' % runlang)
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
                + "><w:body><w:p><w:r>" + rpr + "<w:t>" + TEXT + "</w:t></w:r></w:p>"
                '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
                "</w:sectPr></w:body></w:document>")
    for label, ea, scripts, lang, runlang in ARMS:
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings_for(lang))
            z.writestr("word/theme/theme1.xml", theme_xml(ea, scripts))
            z.writestr("word/document.xml", doc_for(runlang))
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def _shape(ea, scripts):
    faces = ",".join("%s=%s" % kv for kv in sorted((scripts or {}).items())) or "(none)"
    return ("(omitted)" if ea is None else repr(ea)), faces


def pdf():
    import fitz
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    print("%-26s %-9s %-22s %-8s %-7s %s"
          % ("arm", "<a:ea>", "scripts", "themeLg", "runLg", "Word paints"))
    try:
        for label, ea, scripts, lang, runlang in ARMS:
            src, out = docx(label), docx(label).replace(".docx", ".pdf")
            d = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
            try:
                d.ExportAsFixedFormat(out, 17)
                com = str(d.Paragraphs(1).Range.Font.NameFarEast)
            finally:
                d.Close(False)
            faces = sorted({e[3].split("+")[-1] for e in fitz.open(out)[0].get_fonts(False)})
            a, b = _shape(ea, scripts)
            lg = "-" if lang is None else lang.split('w:eastAsia="')[-1].rstrip('"')                 if "eastAsia" in (lang or "") else "en only"
            print("%-26s %-9s %-22s %-8s %-7s %s   [COM=%r]"
                  % (label, a, b, lg, runlang or "-", ", ".join(faces), com))
    finally:
        app.Quit()


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    print("%-26s %-9s %-22s %-8s %-7s %s"
          % ("arm", "<a:ea>", "scripts", "themeLg", "runLg", "Oxi resolves"))
    for label, ea, scripts, lang, runlang in ARMS:
        out = os.path.join(tempfile.gettempdir(), "themeea_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "te"),
                        "--dump-layout=" + out], check=True, capture_output=True, env=env)
        pages = json.load(open(out, encoding="utf-8"))["pages"]
        # The dump carries no family name, so read the quantity the choice
        # MOVES: the line box height. MS Mincho 10.5 -> 13.6, Meiryo -> 20.4,
        # Yu Mincho -> its own. A face that changes nothing here would not
        # change the layout either.
        fams = ["h=%s" % ", ".join(sorted({"%.2f" % e["h"] for p in pages
                                           for e in p["elements"] if e["type"] == "text"}))]
        a, b = _shape(ea, scripts)
        lg = "-" if lang is None else lang.split('w:eastAsia="')[-1].rstrip('"')             if "eastAsia" in (lang or "") else "en only"
        print("%-26s %-9s %-22s %-8s %-7s %s" % (label, a, b, lg, runlang or "-", ", ".join(fams)))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
