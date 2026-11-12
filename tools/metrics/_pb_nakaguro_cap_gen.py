# -*- coding: utf-8 -*-
"""How much overflow does Word absorb on a justified line holding ONE middle dot and one ASCII space?

blind-J ja/reference__6c5dcb5663cdd5fb (compat 15, compressPunctuation, docGrid lines, jc both, 2-column
section, HG丸ｺﾞｼｯｸM-PRO 12): «2020-2021 年度　4R2Z　ゾーン・チェアパーソン» is one line in Word (・ drawn
at 6.0, the space after 2021 at 2.26 of 2.98) with ~6.5pt of overflow; Oxi's S1490 capacity for one mark
(k/(k+1) cell - 0.045 = 5.46) plus the S825 space credit (0.75) wraps the last ン.
Faithful slice: the document itself, w:right swept on that paragraph (and its twin «2021-2022 …»).
Arms: rNNN (twips of right indent), nsNNN (autoSpaceDE/DN off: the «2.98» gap after «2021» is the
CJK/Latin auto-space, not a character -> the ・ alone).
Read: does the paragraph stay one line (the PDF line holding «2020-2021»/«2020-2021年度» ends in «ン»).

    python _pb_nakaguro_cap_gen.py gen | word | oxi EXE
"""
import json, os, subprocess, sys, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_nakaguro_cap")
SRC = r"C:/Users/ryuji/oxi/pipeline_data/docx_corpus/ja/reference/6c5dcb5663cdd5fb.docx"
ANCHOR = '<w:t>4R2Z　ゾーン・チェアパーソン</w:t>'
ARMS = ["r%d" % r for r in range(0, 42, 3)] + ["ns%d" % r for r in range(0, 42, 3)] + ["ns%d" % r for r in range(42, 140, 4)]


def body(arm):
    s = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf8")
    k = s.index(ANCHOR)          # the first paragraph (2020-2021)
    p0 = s.rfind("<w:p ", 0, k)
    p1 = s.index("</w:p>", k)
    para = s[p0:p1]
    r = int(arm.lstrip("ns").lstrip("r"))
    assert '<w:ind w:leftChars="45" w:left="94"/>' in para
    para = para.replace('<w:ind w:leftChars="45" w:left="94"/>', '<w:ind w:leftChars="45" w:left="94" w:right="%d"/>' % r)
    if arm.startswith("ns"):
        para = para.replace('<w:ind w:leftChars', '<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:ind w:leftChars', 1)
    return s[:p0] + para + s[p1:]


def gen():
    os.makedirs(OUT, exist_ok=True)
    z = zipfile.ZipFile(SRC)
    for a in ARMS:
        with zipfile.ZipFile(os.path.join(OUT, a + ".docx"), "w", zipfile.ZIP_DEFLATED) as o:
            for n in z.namelist():
                o.writestr(n, body(a).encode("utf8") if n == "word/document.xml" else z.read(n))
    print("wrote", len(ARMS))


def one_line(pdf):
    import fitz
    for b in fitz.open(pdf)[0].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).replace(" ", "").replace("\u3000", "")
            if t.startswith("2020-2021"):
                return t.endswith("チェアパーソン"), round(l["bbox"][2], 2)
    return None


def word():
    import win32com.client as w
    res = {}
    app = w.DispatchEx("Word.Application"); app.Visible = False
    try:
        for a in ARMS:
            if "--from" in sys.argv and ARMS.index(a) < ARMS.index(sys.argv[sys.argv.index("--from") + 1]):
                continue
            d = app.Documents.Open(os.path.join(OUT, a + ".docx"), ReadOnly=True)
            d.ExportAsFixedFormat(os.path.join(OUT, a + "_word.pdf"), 17); d.Close(False)
            res[a] = one_line(os.path.join(OUT, a + "_word.pdf")); print(a, res[a], flush=True)
    finally:
        app.Quit()
    old = json.load(open(os.path.join(OUT, "word.json"))) if os.path.exists(os.path.join(OUT, "word.json")) else {}
    old.update(res)
    json.dump(old, open(os.path.join(OUT, "word.json"), "w"), indent=1)


def oxi(exe):
    want = json.load(open(os.path.join(OUT, "word.json")))
    for a in ARMS:
        pdf = os.path.join(OUT, a + "_oxi.pdf")
        subprocess.run([exe, "docx-to-pdf", os.path.join(OUT, a + ".docx"), pdf], capture_output=True,
                       env=dict(os.environ, OXI_IGNORE_LRPB="1"))
        print(a, "word", want[a], "oxi", one_line(pdf))


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
