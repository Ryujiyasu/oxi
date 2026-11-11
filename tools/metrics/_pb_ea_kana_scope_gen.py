# -*- coding: utf-8 -*-
"""Han characters under a Latin eastAsia face: does kana nearby change Word's fallback face?

`_pb_ea_latin_face_fallback_gen.py`: Han-only Chinese text under Calibri / Arial / ...
draws in PMingLiU (Times New Roman: Microsoft JhengHei), theme or not, whatever the
lang -- while S1370 / S1594 measured Japanese text under the same faces in the theme's
Jpan face / 游明朝 / ＭＳ 明朝.  Arms (Calibri docDefaults, theme with Jpan ＭＳ 明朝):
  han         Han-only paragraph
  run_kana    Han then «の» in the same run
  para_kana   a Han-only paragraph after a kana paragraph
  runs_mixed  Han-only run, then a kana run, in one paragraph
  kana_first  kana run, then a Han-only run
Read: the PDF font of every span, in order.

    python _pb_ea_kana_scope_gen.py gen / word / oxi EXE
"""
import json
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_ea_latin_face_fallback_gen as base  # noqa: E402

OUT = os.path.join(base.REPO, "pipeline_data", "_pb_ea_kana_scope")
HAN = "报告讲座主题为连续谱中的束缚态涉及量子力学"
KANA = "これはひらがなとカタカナのぶんです"
R = '<w:r><w:t>%s</w:t></w:r>'
ARMS = {
    "han": '<w:p>%s</w:p>' % (R % HAN),
    "run_kana": '<w:p>%s</w:p>' % (R % (HAN + "の")),
    "para_kana": '<w:p>%s</w:p><w:p>%s</w:p>' % (R % KANA, R % HAN),
    "runs_mixed": '<w:p>%s%s</w:p>' % (R % HAN, R % KANA),
    "kana_first": '<w:p>%s%s</w:p>' % (R % KANA, R % HAN),
}


def gen():
    os.makedirs(OUT, exist_ok=True)
    src = os.path.join(base.OUT, "Calibri_nolang_theme.docx")
    z = zipfile.ZipFile(src)
    for n, body in ARMS.items():
        doc = z.read("word/document.xml").decode("utf8")
        i, j = doc.index("<w:body>") + 8, doc.index("<w:sectPr")
        doc = doc[:i] + body + doc[j:]
        with zipfile.ZipFile(os.path.join(OUT, n + ".docx"), "w", zipfile.ZIP_DEFLATED) as o:
            for it in z.infolist():
                o.writestr(it.filename, doc.encode("utf8") if it.filename == "word/document.xml" else z.read(it.filename))
    print("wrote", len(ARMS))


def spans(pdf):
    import fitz
    out = []
    for b in fitz.open(pdf)[0].get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            for s in ln["spans"]:
                if s["text"].strip():
                    out.append((s["font"], s["text"][:6]))
    return out


def word():
    import win32com.client as w
    res = {}
    for n in ARMS:
        app = w.DispatchEx("Word.Application")
        app.Visible = False
        try:
            d = app.Documents.Open(os.path.join(OUT, n + ".docx"), ReadOnly=True)
            d.ExportAsFixedFormat(os.path.join(OUT, n + "_word.pdf"), 17)
            d.Close(False)
        finally:
            app.Quit()
        res[n] = spans(os.path.join(OUT, n + "_word.pdf"))
        print(n, res[n], flush=True)
    json.dump(res, open(os.path.join(OUT, "word.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def oxi(exe):
    want = json.load(open(os.path.join(OUT, "word.json"), encoding="utf-8"))
    for n in ARMS:
        pdf = os.path.join(OUT, n + "_oxi.pdf")
        subprocess.run([exe, "docx-to-pdf", os.path.join(OUT, n + ".docx"), pdf], capture_output=True)
        print(n, "\n   word", want[n], "\n   oxi ", spans(pdf))


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
