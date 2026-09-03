# -*- coding: utf-8 -*-
"""Does a run that paints NOTHING still raise the line?

S1297 (the theme's `<a:font script=...>` entry winning over an empty `<a:ea>`)
is measured and correct, but turning it on costs SSIM on three
kyodokenkyuyoushiki documents. In each of them the only run that resolves
through the theme holds nothing but IDEOGRAPHIC SPACES, and Word's own PDF puts
that line at 14.28pt -- the ＭＳ 明朝 line for 11.04pt text, not the 18.4 a Yu
Mincho run would ask for. So either Word resolves that run to ＭＳ 明朝 after
all, or a run with no visible glyph does not get to raise the line.

The face cannot be read off the PDF (an invisible run embeds no glyph), so this
measures the LINE instead, by difference: every arm is

    [marker]  [arm paragraph]  [marker]

and the arm's height is the marker-to-marker distance minus the same distance
in the `control` arm, which has no paragraph in between.

    python _pb_themerun_gen.py gen
    python _pb_themerun_gen.py pdf      # Word truth
    python _pb_themerun_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_themerun")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402
from _pb_themeea_gen import theme_xml  # noqa: E402

IDEO = "　　"
LIT = "ＭＳ 明朝"

# (label, [(source, text, lang), ...]) — `source` is "theme" (rFonts pointing at
# minorEastAsia, as d1e8ac8's run does) or "literal" (an explicit ＭＳ 明朝).
ARMS = [
    ("control", None),
    ("theme_visible", [("theme", "本文", None)]),
    ("theme_spaces", [("theme", IDEO, None)]),
    ("theme_spaces_zhTW", [("theme", IDEO, "zh-TW")]),
    ("theme_spaces_then_literal", [("theme", IDEO, None), ("literal", "殿", None)]),
    ("theme_spaces_then_literal_zhTW", [("theme", IDEO, "zh-TW"), ("literal", "殿", None)]),
    ("theme_visible_then_literal", [("theme", "本文", None), ("literal", "殿", None)]),
    ("literal_spaces_then_literal", [("literal", IDEO, None), ("literal", "殿", None)]),
]
SZ = 22          # half-points = 11pt, as in d1e8ac8


def docx(label):
    return os.path.join(OUT, "themerun_%s.docx" % label)


def run_xml(source, text, lang):
    if source == "theme":
        # d1e8ac8's shape exactly: ascii/hAnsi/eastAsia all pointing at the EA theme.
        fonts = ('<w:rFonts w:asciiTheme="minorEastAsia" w:eastAsiaTheme="minorEastAsia"'
                 ' w:hAnsiTheme="minorEastAsia" w:hint="eastAsia"/>')
    else:
        fonts = '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:hint="eastAsia"/>' % (LIT, LIT)
    lg = "" if lang is None else '<w:lang w:eastAsia="%s"/>' % lang
    return ('<w:r><w:rPr>%s<w:sz w:val="%d"/><w:szCs w:val="%d"/>%s</w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r>' % (fonts, SZ, SZ, lg, text))


def marker():
    return ('<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr>'
            + run_xml("literal", "基準", None) + "</w:p>")


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
             + rel % (3, "settings", "settings.xml") + "</Relationships>")
    # docDefaults names ＭＳ 明朝 outright, exactly as d1e8ac8 does, so only a run
    # that asks for the theme by name can ever reach it.
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Century" w:eastAsia="%s" w:hAnsi="Century"/>'
              '<w:sz w:val="%d"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/></w:style></w:styles>' % (LIT, SZ))
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                '<w:compat><w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>')
    theme = theme_xml("", {"Jpan": "游明朝"})
    for label, runs in ARMS:
        mid = ""
        if runs:
            mid = ('<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr>'
                   + "".join(run_xml(*r) for r in runs) + "</w:p>")
        body = marker() + mid + marker()
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
                 "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/theme/theme1.xml", theme)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def report(spans_by_arm):
    base = None
    print("%-32s %-8s %-8s %s" % ("arm", "gap", "height", "what the arm holds"))
    for label, ys in spans_by_arm:
        marks = [y for y, t in ys if "基準" in t]
        gap = (max(marks) - min(marks)) if len(marks) >= 2 else None
        if label == "control":
            base = gap
        h = "" if (gap is None or base is None) else "%.2f" % (gap - base)
        desc = dict(ARMS).get(label)
        desc = "(no paragraph)" if not desc else ", ".join(
            "%s:%r%s" % (s, t, "" if lg is None else " lang=" + lg) for s, t, lg in desc)
        print("%-32s %-8s %-8s %s"
              % (label, "-" if gap is None else "%.2f" % gap, h, desc))


def pdf():
    import fitz
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    out = []
    try:
        for label, _ in ARMS:
            src, dst = docx(label), docx(label).replace(".docx", ".pdf")
            d = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
            try:
                d.ExportAsFixedFormat(dst, 17)
            finally:
                d.Close(False)
            ys = []
            for b in fitz.open(dst)[0].get_text("dict")["blocks"]:
                for ln in b.get("lines", []):
                    t = "".join(s["text"] for s in ln["spans"])
                    if t.strip():
                        ys.append((min(s["bbox"][1] for s in ln["spans"]), t))
            out.append((label, ys))
    finally:
        app.Quit()
    print("== WORD ==")
    report(out)


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = []
    for label, _ in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "themerun_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "tr"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        # ★Oxi emits ONE element per character, so a line has to be rebuilt by
        # y before its text can be matched -- reading elements one at a time
        # finds no marker at all and every arm reports "-".
        by_y = {}
        for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
            for e in pg["elements"]:
                if e["type"] == "text" and (e.get("text") or "").strip():
                    by_y.setdefault(round(e["y"], 2), []).append((e["x"], e["text"]))
        ys = [(y, "".join(t for _, t in sorted(v))) for y, v in sorted(by_y.items())]
        out.append((label, ys))
    print("== OXI %s ==" % (envs or "(default)"))
    report(out)


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
