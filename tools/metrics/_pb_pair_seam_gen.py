# -*- coding: utf-8 -*-
"""Does Word compress 」） across a run seam where the two runs name the face differently?

Faithful slice of ja/policies__07543a6b p33 table (the table holding «（ PT「投薬欠落」）»):
the closing ） sits in its own run with no eastAsia slot (resolves to ＭＳ 明朝 via font
linking), 」 in a run naming ＭＳ 明朝.  Arms:
  orig     -- as authored
  explicit -- the ） run also names eastAsia ＭＳ 明朝
  merged   -- ） moved into the 「投薬欠落」 run
Read: lines of that cell paragraph (PDF), the x extent of «「投薬欠落」）».

    python _pb_pair_seam_gen.py gen | word | oxi EXE
"""
import json, os, re, subprocess, sys, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_pair_seam")
SRC = r"C:/Users/ryuji/oxi-main/pipeline_data/docx_corpus/ja/policies/07543a6b9776a1cf.docx"
R_NOEA = '<w:r w:rsidR="00316901" w:rsidRPr="003B1B96"><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr><w:t>）</w:t></w:r></w:p>'
ARMS = ["orig", "explicit", "merged"]
# sweep: right indent (twips) on that paragraph, as authored (orig) -- reads the bill of
# «（ PT「投薬欠落」）» = avail - r at the wrap threshold
SWEEP = ["r%d" % r for r in range(0, 260, 20)]
# fine sweep at the pair threshold: bill 108.67 (」 halved before ）) wraps above 119.6tw,
# 108.73 (） at line end halved, 」 full) above 118.4tw
FINE = ["r%d" % r for r in range(115, 124)]
# solo closer: «（ PT「投薬欠落）» (no pair); ） billed full -> wraps above 238tw, half -> 358tw
SOLO = ["s%d" % r for r in range(200, 400, 20)]
# the pair followed by more text in the same paragraph (亜亜亜 after ）): «」）» ends a LINE, not
# the paragraph; same right indents as the pair sweep's fit (100tw) / wrap (120tw) / 0
TAIL = ["t0", "t100", "t115", "t120"]


def zread(z, name):
    return z.read(name)


def body(arm):
    s = zread(zipfile.ZipFile(SRC), "word/document.xml").decode("utf8")
    j = s.index("<w:t>「投薬欠落」</w:t></w:r>" + R_NOEA)
    t0 = s.rfind("<w:tbl>", 0, j)
    t1 = s.index("</w:tbl>", j) + 8
    tbl = s[t0:t1]
    assert tbl.count(R_NOEA) >= 1, "seam run not found"
    if arm == "explicit":
        tbl = tbl.replace(R_NOEA, R_NOEA.replace('w:hAnsi="Arial" w:cs="Arial"/>', 'w:eastAsia="ＭＳ 明朝" w:hAnsi="Arial" w:cs="Arial"/>'))
    if arm.startswith("t"):
        tbl = tbl.replace(R_NOEA, R_NOEA.replace("<w:t>）</w:t></w:r></w:p>", "<w:t>）亜亜亜</w:t></w:r></w:p>"), 1) if False else tbl
        k0 = tbl.index("<w:t>「投薬欠落」</w:t></w:r>" + R_NOEA)
        k1 = k0 + len("<w:t>「投薬欠落」</w:t></w:r>" + R_NOEA)
        tbl = tbl[:k0] + "<w:t>「投薬欠落」</w:t></w:r>" + R_NOEA.replace("<w:t>）</w:t>", "<w:t>）亜亜亜</w:t>") + tbl[k1:]
        arm = "r" + arm[1:]
    if arm.startswith("s"):
        tbl = tbl.replace("<w:t>「投薬欠落」</w:t></w:r>" + R_NOEA, "<w:t>「投薬欠落</w:t></w:r>" + R_NOEA)
        arm = "r" + arm[1:]
    if arm.startswith("r"):
        a = '<w:ind w:leftChars="20" w:left="42"/>'
        k = tbl.index("<w:t>）亜亜亜</w:t>") - 200 if "）亜亜亜" in tbl else tbl.index(("<w:t>「投薬欠落」</w:t></w:r>" if "「投薬欠落」</w:t></w:r>" + R_NOEA in tbl else "<w:t>「投薬欠落</w:t></w:r>") + R_NOEA)
        p0 = tbl.rfind("<w:p ", 0, k)
        seg = tbl[p0:k]
        assert a in seg
        tbl = tbl[:p0] + seg.replace(a, '<w:ind w:leftChars="20" w:left="42" w:right="%d"/>' % int(arm[1:])) + tbl[k:]
    if arm == "merged":
        tbl = tbl.replace("<w:t>「投薬欠落」</w:t></w:r>" + R_NOEA, "<w:t>「投薬欠落」）</w:t></w:r></w:p>")
        assert "「投薬欠落」）</w:t>" in tbl
    b0 = s.index("<w:body>") + 8
    sp = s.rfind("<w:sectPr")
    return s[:b0] + tbl + "<w:p/>" + s[sp:]


def gen():
    os.makedirs(OUT, exist_ok=True)
    z = zipfile.ZipFile(SRC)
    for a in ARMS + SWEEP + FINE + SOLO + TAIL:
        with zipfile.ZipFile(os.path.join(OUT, a + ".docx"), "w", zipfile.ZIP_DEFLATED) as o:
            for it in z.infolist():
                b = zread(z, it.filename)
                if it.filename == "word/document.xml":
                    b = body(a).encode("utf8")
                o.writestr(it.filename, b)
    print("wrote", ARMS)


def lines(pdf):
    import fitz
    res = []
    for pg in fitz.open(pdf):
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(s["text"] for s in l["spans"])
                if ("PT" in t or "欠落」" in t or "亜" in t or t.strip() in ("落」）", "）", "落）")) and 255 < l["bbox"][0] < 390:
                    res.append((round(l["bbox"][0], 2), round(l["bbox"][2], 2), round(l["bbox"][1], 1), t))
    return res


def word(arms=None):
    import win32com.client as w
    res = {}
    for a in arms or ARMS:
        app = w.DispatchEx("Word.Application"); app.Visible = False
        try:
            d = app.Documents.Open(os.path.join(OUT, a + ".docx"), ReadOnly=True)
            d.ExportAsFixedFormat(os.path.join(OUT, a + "_word.pdf"), 17); d.Close(False)
        finally:
            app.Quit()
        res[a] = lines(os.path.join(OUT, a + "_word.pdf")); print(a, res[a], flush=True)
    json.dump(res, open(os.path.join(OUT, "word%s.json" % (("_sweep%d" % len(arms)) if arms else "")), "w", encoding="utf8"), ensure_ascii=False, indent=1)


def oxi(exe):
    want = json.load(open(os.path.join(OUT, "word.json"), encoding="utf8"))
    for a in ARMS:
        pdf = os.path.join(OUT, a + "_oxi.pdf")
        subprocess.run([exe, "docx-to-pdf", os.path.join(OUT, a + ".docx"), pdf], capture_output=True)
        print(a, "\n  word", want[a], "\n  oxi ", lines(pdf))


if __name__ == "__main__":
    {"gen": gen, "word": word, "sweep": lambda: word(SWEEP), "fine": lambda: word(FINE + SOLO), "tail": lambda: word(TAIL)}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
