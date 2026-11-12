# -*- coding: utf-8 -*-
"""Does a modern (compat 15, doNotCompress) table cell halve a closing-mark pair, and is w:kern the gate?

Slice of ja/policies__07543a6b (the p33 table, «（ PT「投薬欠落」）» cell, compat 15,
doNotCompress, Normal kern 2).  The cell paragraph's text becomes a pair sampler
«投」）落、「投。」落）」投）、落» (one run, eastAsia ＭＳ 明朝).  Arms: kern (as authored) /
nokern (Normal's w:kern removed) / body (same paragraph moved out of the table, kern).
Read: per-character advances of the sampler line out of Word's PDF.

    python _pb_cell_pair_kern_gen.py gen | word | oxi EXE
"""
import json, os, subprocess, sys, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cell_pair_kern")
SRC = os.path.join(REPO, "pipeline_data", "_pb_pair_seam", "merged.docx")
SAMPLER = "投」）落、「投。」落）」投）、落"
ARMS = ["kern", "nokern", "body", "body_nokern", "body_nokern_en", "body_kern_en"]


def gen():
    os.makedirs(OUT, exist_ok=True)
    z = zipfile.ZipFile(SRC)
    doc = z.read("word/document.xml").decode("utf8")
    assert doc.count("<w:t>「投薬欠落」）</w:t>") == 1
    doc = doc.replace("<w:t>「投薬欠落」）</w:t>", "<w:t>%s</w:t>" % SAMPLER)
    sty = z.read("word/styles.xml").decode("utf8")
    for a in ARMS:
        d, s = doc, sty
        if a.endswith("_en"):
            # Normal's eastAsia language removed -> docDefaults' en-US
            i = s.index('w:styleId="a"')
            j = s.index("</w:style>", i)
            assert '<w:lang w:eastAsia="ja-JP"/>' in s[i:j]
            s = s[:i] + s[i:j].replace('<w:lang w:eastAsia="ja-JP"/>', "") + s[j:]
        if a in ("nokern", "body_nokern", "body_nokern_en"):
            i = s.index('w:styleId="a"')
            j = s.index("</w:style>", i)
            assert '<w:kern w:val="2"/>' in s[i:j]
            s = s[:i] + s[i:j].replace('<w:kern w:val="2"/>', "") + s[j:]
        if a.startswith("body"):
            k = d.index(SAMPLER)
            p0, p1 = d.rfind("<w:p ", 0, k), d.index("</w:p>", k) + 6
            b0, sp = d.index("<w:body>") + 8, d.rfind("<w:sectPr")
            d = d[:b0] + d[p0:p1] + d[sp:]
        with zipfile.ZipFile(os.path.join(OUT, a + ".docx"), "w", zipfile.ZIP_DEFLATED) as o:
            for n in z.namelist():
                b = z.read(n)
                if n == "word/document.xml":
                    b = d.encode("utf8")
                if n == "word/styles.xml":
                    b = s.encode("utf8")
                o.writestr(n, b)
    print("wrote", ARMS)


CELL = False


def adv(pdf):
    """Within-line pair advances of the sampler, over however many lines it wraps to."""
    import fitz
    lines = []
    for pg in fitz.open(pdf):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                # the sampler's cell column is x 259.6..387.2 (gridCol 2551tw); the body arms
                # run from the left margin, so take the column only when the text is in it
                chs = [(c["c"], c["bbox"][0]) for s in l["spans"] for c in s["chars"] if c["c"].strip()
                       and (not CELL or 255 < c["bbox"][0] < 388)]
                if chs:
                    lines.append((round(l["bbox"][1], 1), chs))
    lines.sort(key=lambda x: x[0])
    # other cells of the slice share the column band: follow the sampler line by line,
    # taking only lines that continue it
    out, rest = [], None
    for _, chs in lines:
        t = "".join(c[0] for c in chs)
        if rest is None:
            if SAMPLER[:3] not in t:
                continue
            chs = chs[t.index(SAMPLER[:3]):]
            rest = SAMPLER
        elif not rest.startswith(t[:len(rest)][:2]):
            continue
        seg = chs[:len(rest)]
        if "".join(c[0] for c in seg) != rest[:len(seg)]:
            continue
        out += [(seg[k][0] + seg[k + 1][0], round(seg[k + 1][1] - seg[k][1], 2)) for k in range(len(seg) - 1)]
        rest = rest[len(seg):]
        if not rest:
            return out
    return out or None


def word():
    global CELL
    import win32com.client as w
    res = {}
    for a in ARMS:
        if os.path.exists(os.path.join(OUT, a + "_word.pdf")) and "--reuse" in sys.argv:
            CELL = not a.startswith("body")
            res[a] = adv(os.path.join(OUT, a + "_word.pdf")); print(a, res[a]); continue
        app = w.DispatchEx("Word.Application"); app.Visible = False
        try:
            d = app.Documents.Open(os.path.join(OUT, a + ".docx"), ReadOnly=True)
            d.ExportAsFixedFormat(os.path.join(OUT, a + "_word.pdf"), 17); d.Close(False)
        finally:
            app.Quit()
        CELL = not a.startswith("body")
        res[a] = adv(os.path.join(OUT, a + "_word.pdf")); print(a, res[a], flush=True)
    json.dump(res, open(os.path.join(OUT, "word.json"), "w", encoding="utf8"), ensure_ascii=False, indent=1)


def oxi(exe):
    global CELL
    want = json.load(open(os.path.join(OUT, "word.json"), encoding="utf8"))
    for a in ARMS:
        pdf = os.path.join(OUT, a + "_oxi.pdf")
        subprocess.run([exe, "docx-to-pdf", os.path.join(OUT, a + ".docx"), pdf], capture_output=True)
        CELL = not a.startswith("body")
        got = adv(pdf)
        diff = [(w_[0], w_[1], o_[1]) for w_, o_ in zip(want[a] or [], got or []) if abs(w_[1] - o_[1]) > 0.05]
        print(a, "match" if not diff and got else diff or got)


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
