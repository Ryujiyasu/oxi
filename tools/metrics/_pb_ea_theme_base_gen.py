# -*- coding: utf-8 -*-
"""Which face carries CJK text whose eastAsia slot names a Latin face, by theme Jpan face?

`_pb_cjk_glyph_fallback2_gen.py` found the base face MS Mincho under a theme Jpan of ＭＳ 明朝,
游明朝 and no theme alike (docDefaults rFonts eastAsia="Calibri"); S1594 (old truth machine)
had the theme's Jpan face win.  Arms: docDefaults eastAsia {literal Calibri, eastAsiaTheme
minorEastAsia} x theme Jpan {ＭＳ 明朝, 游明朝, ＭＳ Ｐゴシック, メイリオ, none}.  Text: kana + common
kanji («あいう日本語の文章»).  Read: the PDF span font of that text.

    python _pb_ea_theme_base_gen.py gen / word / oxi EXE
"""
import json
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_ea_latin_face_fallback_gen as base  # noqa: E402

OUT = os.path.join(base.REPO, "pipeline_data", "_pb_ea_theme_base")
JPAN = {"mincho": "ＭＳ 明朝", "yumin": "游明朝", "pgothic": "ＭＳ Ｐゴシック", "meiryo": "メイリオ", "none": None}
SLOTS = {"literal": 'w:eastAsia="Calibri"', "theme": 'w:eastAsiaTheme="minorEastAsia"'}
TEXT = "あいう日本語の文章"


def arms():
    return [(s, j) for s in SLOTS for j in JPAN]


def gen():
    os.makedirs(OUT, exist_ok=True)
    src = zipfile.ZipFile(os.path.join(base.OUT, "Calibri_nolang_theme.docx"))
    for s, j in arms():
        face = JPAN[j]
        with zipfile.ZipFile(os.path.join(OUT, "%s_%s.docx" % (s, j)), "w", zipfile.ZIP_DEFLATED) as o:
            for it in src.infolist():
                b = src.read(it.filename)
                if it.filename == "word/document.xml":
                    d = b.decode("utf8")
                    i, k = d.index("<w:body>") + 8, d.index("<w:sectPr")
                    b = (d[:i] + '<w:p><w:r><w:t>%s</w:t></w:r></w:p>' % TEXT + d[k:]).encode("utf8")
                if it.filename == "word/styles.xml":
                    b = b.decode("utf8").replace('w:eastAsia="Calibri"', SLOTS[s]).encode("utf8")
                if it.filename == "word/theme/theme1.xml":
                    if face is None:
                        continue
                    b = b.decode("utf8").replace("ＭＳ 明朝", face).encode("utf8")
                if face is None and it.filename in ("[Content_Types].xml", "word/_rels/document.xml.rels"):
                    t = b.decode("utf8")
                    t = t.replace('<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>', "")
                    t = t.replace('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>', "")
                    b = t.encode("utf8")
                o.writestr(it.filename, b)
    print("wrote", len(arms()))


def fonts(pdf):
    import fitz
    return sorted({s["font"] for b in fitz.open(pdf)[0].get_text("dict")["blocks"] for l in b.get("lines", [])
                   for s in l["spans"] if s["text"].strip()})


def word():
    import win32com.client as w
    res = {}
    for s, j in arms():
        n = "%s_%s" % (s, j)
        app = w.DispatchEx("Word.Application")
        app.Visible = False
        try:
            d = app.Documents.Open(os.path.join(OUT, n + ".docx"), ReadOnly=True)
            d.ExportAsFixedFormat(os.path.join(OUT, n + "_word.pdf"), 17)
            d.Close(False)
        finally:
            app.Quit()
        res[n] = fonts(os.path.join(OUT, n + "_word.pdf"))
        print(n, res[n], flush=True)
    json.dump(res, open(os.path.join(OUT, "word.json"), "w"), indent=1)


def oxi(exe):
    want = json.load(open(os.path.join(OUT, "word.json")))
    for s, j in arms():
        n = "%s_%s" % (s, j)
        pdf = os.path.join(OUT, n + "_oxi.pdf")
        subprocess.run([exe, "docx-to-pdf", os.path.join(OUT, n + ".docx"), pdf], capture_output=True)
        print("%-16s word %-28s oxi %s" % (n, want[n], fonts(pdf)))


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
