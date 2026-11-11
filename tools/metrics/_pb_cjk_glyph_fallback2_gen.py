# -*- coding: utf-8 -*-
"""Which face draws a Han character the CJK base face lacks?  (re-measure of S1591)

S1591 (2026-09-27, the old truth machine): MS Mincho-less characters go to Yu Gothic
when it has them, else SimSun, PMingLiU only for Big5 characters.  On the current
Word, zh dev educational__005919f8's 报 / 讲 / 连续谱 (not Big5) draw in PMingLiU.
Characters MS Mincho lacks, by class (Yu Gothic covers? / PMingLiU covers? / in Big5?):
  ming     -/ming/-      e.g. 专业丛东
  ming_b5  -/ming/b5     e.g. 丳乒乓
  yug      yug/ming/-    e.g. 丷伫兴兹
  yug_b5   yug/ming/b5   e.g. 俓偷僰
  sim      -/-/-         (SimSun only) e.g. 龵龶
Arms: Calibri docDefaults eastAsia x theme Jpan ＭＳ 明朝 / 游明朝 / no theme.
Read: the PDF span font of each character.

    python _pb_cjk_glyph_fallback2_gen.py gen / word / oxi EXE
"""
import json
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_ea_latin_face_fallback_gen as base  # noqa: E402

OUT = os.path.join(base.REPO, "pipeline_data", "_pb_cjk_glyph_fallback2")
CLASSES = {"base": "国中的人大日本学", "ming": "专业丛东丝丢严丧", "ming_b5": "丳乒乓亃亄亓仉仜", "yug": "丷伫兴兹凟刦别厉",
           "yug_b5": "俓偷僰內凔吳唧垮", "sim": "龵龶龷龸龹龺龻鿍"}
THEMES = {"mincho": "ＭＳ 明朝", "yumin": "游明朝", "none": None}


def gen():
    os.makedirs(OUT, exist_ok=True)
    src = zipfile.ZipFile(os.path.join(base.OUT, "Calibri_nolang_theme.docx"))
    for tn, face in THEMES.items():
        body = "".join('<w:p><w:r><w:t>%s:</w:t></w:r><w:r><w:t>%s</w:t></w:r></w:p>' % (k, v) for k, v in CLASSES.items())
        doc = src.read("word/document.xml").decode("utf8")
        i, j = doc.index("<w:body>") + 8, doc.index("<w:sectPr")
        doc = doc[:i] + body + doc[j:]
        with zipfile.ZipFile(os.path.join(OUT, tn + ".docx"), "w", zipfile.ZIP_DEFLATED) as o:
            for it in src.infolist():
                b = src.read(it.filename)
                if it.filename == "word/document.xml":
                    b = doc.encode("utf8")
                if it.filename == "word/theme/theme1.xml":
                    if face is None:
                        continue
                    b = b.decode("utf8").replace("ＭＳ 明朝", face).encode("utf8")
                if face is None and it.filename in ("[Content_Types].xml", "word/_rels/document.xml.rels"):
                    s = b.decode("utf8")
                    s = s.replace('<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>', "")
                    s = s.replace('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>', "")
                    b = s.encode("utf8")
                o.writestr(it.filename, b)
    print("wrote", len(THEMES))


def per_char(pdf):
    import fitz
    out = {}
    for b in fitz.open(pdf)[0].get_text("rawdict")["blocks"]:
        for ln in b.get("lines", []):
            for s in ln["spans"]:
                for c in s["chars"]:
                    if ord(c["c"]) >= 0x3040:
                        out[c["c"]] = s["font"]
    return out


def table(fonts):
    return {k: [fonts.get(ch, "?") for ch in v] for k, v in CLASSES.items()}


def word():
    import win32com.client as w
    res = {}
    for tn in THEMES:
        app = w.DispatchEx("Word.Application")
        app.Visible = False
        try:
            d = app.Documents.Open(os.path.join(OUT, tn + ".docx"), ReadOnly=True)
            d.ExportAsFixedFormat(os.path.join(OUT, tn + "_word.pdf"), 17)
            d.Close(False)
        finally:
            app.Quit()
        res[tn] = table(per_char(os.path.join(OUT, tn + "_word.pdf")))
        print(tn, json.dumps(res[tn], ensure_ascii=False), flush=True)
    json.dump(res, open(os.path.join(OUT, "word.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def oxi(exe):
    want = json.load(open(os.path.join(OUT, "word.json"), encoding="utf-8"))
    bad = 0
    for tn in THEMES:
        pdf = os.path.join(OUT, tn + "_oxi.pdf")
        subprocess.run([exe, "docx-to-pdf", os.path.join(OUT, tn + ".docx"), pdf], capture_output=True)
        got = table(per_char(pdf))
        for k in CLASSES:
            if got[k] != want[tn][k]:
                bad += 1
                print(tn, k, "word", want[tn][k], "oxi", got[k])
    print("mismatching classes", bad)


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
