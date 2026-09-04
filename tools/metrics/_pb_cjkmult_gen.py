# -*- coding: utf-8 -*-
"""What does a CJK line's spacing MULTIPLIER multiply?

`d1e8ac8`'s heading is 14pt CJK with `<w:spacing w:line="360" w:lineRule="auto"/>`
(x1.5). Word advances 24.00pt over it; Oxi advances 27.50 = ＭＳ 明朝's 83/64
natural (18.156) x 1.5 = 27.24. Word's 24.00 divided by 1.5 is 16.0, which is
neither the 83/64 natural nor the hhea one (14.0) -- but it IS about Century's
natural at 14pt (16.38), and Century is what this document's docDefaults name
for `w:ascii` while `w:eastAsia` is ＭＳ 明朝.

So the question is whether the ASCII font takes part in a CJK line's height when
a multiplier is in play. The discriminating arm is the pair that differs ONLY in
`w:ascii`: if the pitch moves, it does.

Word's reported positions are quantised to the 0.75pt device step, so a single
gap cannot resolve 0.3pt. Each arm is therefore ONE paragraph that wraps ~40
times and the pitch is read as (last - first) / (n - 1), the technique
`_pb_linepitch_gen.py` uses -- the quantisation is spent once over 39 gaps.

    python _pb_cjkmult_gen.py gen
    python _pb_cjkmult_gen.py pdf      # Word truth
    python _pb_cjkmult_gen.py oxi      # Oxi, same arms
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cjkmult")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

MINCHO = "ＭＳ 明朝"
BODY = "本文の行送りを測るための文字列です。" * 24

# (label, half-points, w:line value, docDefaults ascii face)
ARMS = [
    ("sz21_x1.0_century", 21, 240, "Century"),
    ("sz21_x1.5_century", 21, 360, "Century"),
    ("sz28_x1.0_century", 28, 240, "Century"),
    ("sz28_x1.5_century", 28, 360, "Century"),      # d1e8ac8's heading shape
    ("sz28_x2.0_century", 28, 480, "Century"),
    # ★The discriminator: same CJK line, different ASCII face.
    ("sz28_x1.5_mincho", 28, 360, MINCHO),
    ("sz28_x1.5_arial", 28, 360, "Arial"),
    ("sz28_x1.5_meiryo", 28, 360, "メイリオ"),
]


def docx(label):
    return os.path.join(OUT, "cjkmult_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/settings.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                    "</Types>")
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
             'relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
             'officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
             'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             "</Relationships>")
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                '<w:compat><w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>')
    for label, sz, line, ascii_face in ARMS:
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  "<w:docDefaults><w:rPrDefault><w:rPr>"
                  '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
                  '<w:sz w:val="%d"/></w:rPr></w:rPrDefault>'
                  "<w:pPrDefault><w:pPr>"
                  '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
                  "</w:pPr></w:pPrDefault></w:docDefaults>"
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
                  '<w:name w:val="Normal"/></w:style></w:styles>'
                  % (ascii_face, MINCHO, ascii_face, sz))
        # One long paragraph, no grid: the pitch is (last - first) / (n - 1).
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body>"
               + '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d"'
                 ' w:lineRule="auto"/></w:pPr><w:r><w:rPr>'
                 '<w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr>'
                 "<w:t>%s</w:t></w:r></w:p>" % (line, sz, BODY)
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
                 "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def report(pitches, who):
    print("== %s ==" % who)
    print("%-22s %-6s %-6s %-10s %-8s %s"
          % ("arm", "pt", "mult", "ascii", "pitch", "pitch / mult"))
    for label, sz, line, ascii_face in ARMS:
        p = pitches.get(label)
        mult = line / 240.0
        print("%-22s %-6.1f %-6.2f %-10s %-8s %s"
              % (label, sz / 2.0, mult, ascii_face,
                 "-" if p is None else "%.3f" % p,
                 "-" if p is None else "%.3f" % (p / mult)))


def pdf():
    import fitz
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    out = {}
    try:
        for label, _, _, _ in ARMS:
            src, dst = docx(label), docx(label).replace(".docx", ".pdf")
            d = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
            try:
                d.ExportAsFixedFormat(dst, 17)
            finally:
                d.Close(False)
            ys = sorted({round(s["bbox"][1], 2)
                         for b in fitz.open(dst)[0].get_text("dict")["blocks"]
                         for ln in b.get("lines", []) for s in ln["spans"]
                         if s["text"].strip()})
            if len(ys) > 3:
                out[label] = (ys[-1] - ys[0]) / (len(ys) - 1)
    finally:
        app.Quit()
    report(out, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = {}
    for label, _, _, _ in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "cjkmult_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "cm"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        ys = sorted({round(e["y"], 2)
                     for pg in json.load(open(dump, encoding="utf-8"))["pages"]
                     for e in pg["elements"]
                     if e["type"] == "text" and (e.get("text") or "").strip()})
        if len(ys) > 3:
            out[label] = (ys[-1] - ys[0]) / (len(ys) - 1)
    report(out, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
