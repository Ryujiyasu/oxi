# -*- coding: utf-8 -*-
"""Is w:kern the gate of closing-mark pair halving inside a table cell, as S547 found in the body?

S547's matrix docs (ＭＳ 明朝 10.5 everywhere, no theme, no lang) re-measured by PDF hold on the
current Word: kern 0 never halves, kern 2 halves 、（ and （「.  The cell wrappers in Oxi halve a
CLOSING mark before another mark only in legacy cells (compat < 15 + compressPunctuation).
Arms: S547 docs (kern 0/2 x compressPunctuation off/on x compat absent/15) with the sampler
«投」）落、「投。」落）」投）、落、（国» placed in a one-cell table (4000tw, wide enough for
one line), and the same in the body.  Read: per-pair advances from Word's PDF.

    python _pb_cell_pair_kern2_gen.py gen | word | oxi EXE
"""
import json, os, subprocess, sys, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cell_pair_kern2")
SRCDIR = r"C:/Users/ryuji/oxi/tools/golden-test/repros/s547_kern"
SAMPLER = "投」）落、「投。」落）」投）、落、（国"
PAIRS = ("」）", "、「", "。」", "）」", "）、", "、（")


def arms():
    return [(k, c, v, w) for k in (0, 1) for c in (0, 1) for v in (0, 1) for w in ("cell", "body")]


def name(a):
    return "k%d_c%d_v%d_%s" % a


def gen():
    os.makedirs(OUT, exist_ok=True)
    for a in arms():
        k, c, v, where = a
        z = zipfile.ZipFile(os.path.join(SRCDIR, "s547b_k%d_c%d_v%d.docx" % (k, c, v)))
        d = z.read("word/document.xml").decode("utf8")
        p = ('<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
             '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % SAMPLER)
        if where == "cell":
            p = ('<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
                 '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
                 + p + '</w:tc></w:tr></w:tbl><w:p/>')
        b0, sp = d.index("<w:body>") + 8, d.index("<w:sectPr")
        d = d[:b0] + p + d[sp:]
        with zipfile.ZipFile(os.path.join(OUT, name(a) + ".docx"), "w", zipfile.ZIP_DEFLATED) as o:
            for n in z.namelist():
                o.writestr(n, d.encode("utf8") if n == "word/document.xml" else z.read(n))
    print("wrote", len(arms()))


def pairs(pdf):
    import fitz
    out = {}
    for pg in fitz.open(pdf):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                ch = [(c["c"], c["bbox"][0]) for s in l["spans"] for c in s["chars"]]
                for i in range(len(ch) - 1):
                    if ch[i][0] + ch[i + 1][0] in PAIRS:
                        out.setdefault(ch[i][0] + ch[i + 1][0], round(ch[i + 1][1] - ch[i][1], 2))
    return out


def word():
    import win32com.client as w
    res = {}
    app = w.DispatchEx("Word.Application"); app.Visible = False
    try:
        for a in arms():
            n = name(a)
            d = app.Documents.Open(os.path.join(OUT, n + ".docx"), ReadOnly=True)
            d.ExportAsFixedFormat(os.path.join(OUT, n + "_word.pdf"), 17); d.Close(False)
            res[n] = pairs(os.path.join(OUT, n + "_word.pdf")); print(n, res[n], flush=True)
    finally:
        app.Quit()
    json.dump(res, open(os.path.join(OUT, "word.json"), "w", encoding="utf8"), ensure_ascii=False, indent=1)


def oxi(exe):
    want = json.load(open(os.path.join(OUT, "word.json"), encoding="utf8"))
    bad = 0
    for a in arms():
        n = name(a)
        pdf = os.path.join(OUT, n + "_oxi.pdf")
        subprocess.run([exe, "docx-to-pdf", os.path.join(OUT, n + ".docx"), pdf], capture_output=True)
        got = pairs(pdf)
        diff = {k: (want[n].get(k), got.get(k)) for k in PAIRS if abs((want[n].get(k) or 0) - (got.get(k) or 0)) > 0.1}
        bad += bool(diff)
        print("%-16s %s" % (n, "match" if not diff else diff))
    print("mismatch arms:", bad)


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
