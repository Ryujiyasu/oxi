# -*- coding: utf-8 -*-
"""Line pitch of PMingLiU text: on its own, and as the fallback inside an MS Mincho line.

OXI_PMINGLIU_ALL put the fallback face right (ja/reports__33da7949: 137 characters in
PMingLiU, as Word), but the line pitch is 18.89 in Oxi against Word's 18.74 at 12pt
(`_pb_ea_latin_face_fallback_gen.py`, Calibri docDefaults), and the page split of that
document moved.  PMingLiU: upm 1024, win 820/204, hhea lineGap 204 (MS Mincho: 0).
Arms x size 10.5 / 12 / 16, single spacing, no docGrid, 4-line paragraphs:
  pm       eastAsia PMingLiU, simplified Han text (all in PMingLiU)
  mix      eastAsia Calibri docDefaults -> MS Mincho base + PMingLiU fallback chars
  mincho   eastAsia MS Mincho, text the base face has (control)
Read: baseline pitch from Word's PDF and Oxi's CLI PDF.

    python _pb_pmingliu_pitch_gen.py gen / word / oxi EXE
"""
import json
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_ea_latin_face_fallback_gen as base  # noqa: E402

OUT = os.path.join(base.REPO, "pipeline_data", "_pb_pmingliu_pitch")
MIXTXT = "报告讲座主题为连续谱中的束缚态涉及量子力学" * 6
MINTXT = "日本語の文章を書いて行の高さを測ります。" * 6
ARMS = {"pm": ("PMingLiU", MIXTXT), "mix": ("Calibri", MIXTXT), "mincho": ("MS Mincho", MINTXT)}
SIZES = [21, 24, 32]


def gen():
    os.makedirs(OUT, exist_ok=True)
    src = zipfile.ZipFile(os.path.join(base.OUT, "Calibri_nolang_theme.docx"))
    for an, (face, txt) in ARMS.items():
        for sz in SIZES:
            body = '<w:p><w:r><w:rPr><w:rFonts w:eastAsia="%s"/><w:sz w:val="%d"/></w:rPr><w:t>%s</w:t></w:r></w:p>' % (face, sz, txt)
            doc = src.read("word/document.xml").decode("utf8")
            i, j = doc.index("<w:body>") + 8, doc.index("<w:sectPr")
            doc = doc[:i] + body + doc[j:]
            with zipfile.ZipFile(os.path.join(OUT, "%s_%d.docx" % (an, sz)), "w", zipfile.ZIP_DEFLATED) as o:
                for it in src.infolist():
                    o.writestr(it.filename, doc.encode("utf8") if it.filename == "word/document.xml" else src.read(it.filename))
    print("wrote", len(ARMS) * len(SIZES))


def pitch(pdf):
    import fitz
    ys = sorted({round(l["spans"][0]["origin"][1], 2) for b in fitz.open(pdf)[0].get_text("dict")["blocks"]
                 for l in b.get("lines", []) if "".join(s["text"] for s in l["spans"]).strip()})
    return round((ys[-1] - ys[0]) / (len(ys) - 1), 3) if len(ys) > 1 else None


def names():
    return ["%s_%d" % (a, s) for a in ARMS for s in SIZES]


def word():
    import win32com.client as w
    res = {}
    for n in names():
        app = w.DispatchEx("Word.Application")
        app.Visible = False
        try:
            d = app.Documents.Open(os.path.join(OUT, n + ".docx"), ReadOnly=True)
            d.ExportAsFixedFormat(os.path.join(OUT, n + "_word.pdf"), 17)
            d.Close(False)
        finally:
            app.Quit()
        res[n] = pitch(os.path.join(OUT, n + "_word.pdf"))
        print(n, res[n], flush=True)
    json.dump(res, open(os.path.join(OUT, "word.json"), "w"), indent=1)


def oxi(exe, envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        env[kv] = "1"
    want = json.load(open(os.path.join(OUT, "word.json")))
    for n in names():
        pdf = os.path.join(OUT, n + "_oxi.pdf")
        subprocess.run([exe, "docx-to-pdf", os.path.join(OUT, n + ".docx"), pdf], capture_output=True, env=env)
        print(n, "word", want[n], "oxi", pitch(pdf))


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else ""))()
