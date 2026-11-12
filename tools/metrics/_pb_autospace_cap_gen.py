# -*- coding: utf-8 -*-
"""Does a justified compat-15 CJK line absorb overflow by shrinking its CJK/Latin auto-space?

_pb_nakaguro_cap_gen.py (blind-J ja/reference__6c5dcb56, one ・ + one «2021|年度» auto-space junction):
Oxi = Word on every arm without auto-space (fits through w:right 42tw, wraps from 46); with the
junction Word still fits at 0 while Oxi wraps -- the junction is worth ~0.5-0.9pt of capacity in Word.
Here, no marks at all: the same paragraph (same slice, HG丸ｺﾞｼｯｸM-PRO 12, jc both inherited, col 2 of
the 2-column section, 288.6pt measure) holds «亜» + «R亜亜» x j + «亜» x k (each R = two junctions),
j = 0..3, with auto-space on (a) and off (n); w:right swept around the expected wrap.
Read: one line or not (the paragraph's first PDF line ends with its last character).

    python _pb_autospace_cap_gen.py gen | word | oxi EXE
"""
import json, os, subprocess, sys, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_autospace_cap")
SRC = r"C:/Users/ryuji/oxi/pipeline_data/docx_corpus/ja/reference/6c5dcb5663cdd5fb.docx"
ANCHOR = '<w:t>4R2Z　ゾーン・チェアパーソン</w:t>'
KS = {0: 23, 1: 20, 2: 17, 3: 14}
# expected wrap (twips of right indent) per (j, autospace): centre of each window
CENTRE = {(0, "a"): 12, (1, "a"): -13, (2, "a"): -38, (3, "a"): -62,
          (0, "n"): 12, (1, "n"): 77, (2, "n"): 142, (3, "n"): 208}


def text(j):
    return "亜" + "R亜亜" * j + "亜" * KS[j]


def arms():
    out = []
    for (j, m), c in CENTRE.items():
        for r in range(c - 44, c + 45, 4):
            out.append("j%d%s_%d" % (j, m, r))
    return out


def parse(a):
    head, r = a.split("_")
    return int(head[1]), head[2], int(r)


def body(a):
    j, m, r = parse(a)
    s = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf8")
    k = s.index(ANCHOR)
    p0 = s.rfind("<w:p ", 0, k)
    p1 = s.index("</w:p>", k)
    para = s[p0:p1]
    ppr_end = para.index("</w:pPr>") + len("</w:pPr>")
    run = para[ppr_end:]
    rpr = run[run.index("<w:rPr>"):run.index("</w:rPr>") + len("</w:rPr>")]
    ind = '<w:ind w:leftChars="45" w:left="94" w:right="%d"/>' % r
    ppr = para[:ppr_end].replace('<w:ind w:leftChars="45" w:left="94"/>', ind)
    assert ind in ppr
    if m == "n":
        ppr = ppr.replace(ind, '<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>' + ind)
    para = ppr + "<w:r>" + rpr + "<w:t>%s</w:t></w:r>" % text(j)
    return s[:p0] + para + s[p1:]


def gen():
    os.makedirs(OUT, exist_ok=True)
    z = zipfile.ZipFile(SRC)
    for a in arms():
        with zipfile.ZipFile(os.path.join(OUT, a + ".docx"), "w", zipfile.ZIP_DEFLATED) as o:
            for n in z.namelist():
                o.writestr(n, body(a).encode("utf8") if n == "word/document.xml" else z.read(n))
    print("wrote", len(arms()))


def one_line(pdf, j):
    import fitz
    t0 = text(j)
    for b in fitz.open(pdf)[0].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).replace(" ", "")
            if t.startswith(t0[:6]):
                return t == t0
    return None


def word():
    import win32com.client as w
    res = {}
    app = w.DispatchEx("Word.Application"); app.Visible = False
    try:
        for a in arms():
            d = app.Documents.Open(os.path.join(OUT, a + ".docx"), ReadOnly=True)
            d.ExportAsFixedFormat(os.path.join(OUT, a + "_word.pdf"), 17); d.Close(False)
            res[a] = one_line(os.path.join(OUT, a + "_word.pdf"), parse(a)[0])
    finally:
        app.Quit()
    json.dump(res, open(os.path.join(OUT, "word.json"), "w"), indent=1)
    report(res)


def report(res):
    for (j, m) in CENTRE:
        row = [(parse(a)[2], v) for a, v in res.items() if parse(a)[:2] == (j, m)]
        row.sort()
        fit = [r for r, v in row if v]
        wrap = [r for r, v in row if v is False]
        print("j=%d %s  fits through r=%s, wraps from r=%s" % (j, m, max(fit) if fit else None,
              min((r for r in wrap if not fit or r > max(fit)), default=None)))


def oxi(exe):
    res = {}
    for a in arms():
        pdf = os.path.join(OUT, a + "_oxi.pdf")
        subprocess.run([exe, "docx-to-pdf", os.path.join(OUT, a + ".docx"), pdf], capture_output=True,
                       env=dict(os.environ, OXI_IGNORE_LRPB="1"))
        res[a] = one_line(pdf, parse(a)[0])
    print("Word:"); report(json.load(open(os.path.join(OUT, "word.json"))))
    print("Oxi:"); report(res)


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
