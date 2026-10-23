# -*- coding: utf-8 -*-
"""Faithful slice: one real paragraph, its document's styles / settings / section, repeated
with the right indent stepped; reports where Word and Oxi move a given character off a line.

    PB_DOC=ja/administrative__0ba21b0c760f2edd PB_NEEDLE=漢字の書き取り PB_X=が PB_OXI=<exe> \
        python tools/metrics/_pb_unit_slice_gen.py

Readout per indent: the text of the line holding the character just before PB_X's first
occurrence after PB_AFTER (default: PB_X itself), i.e. whether PB_X stays on that line.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import feature_census as FC  # noqa: E402

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_unit_slice")
DID = os.environ["PB_DOC"]
NEEDLE = os.environ["PB_NEEDLE"]
CONTEXT = os.environ["PB_CONTEXT"]  # text that ends right before X on the line in question
X = os.environ["PB_X"]
LO, HI, STEP = (int(v) for v in os.environ.get("PB_RANGE", "-300,300,10").split(","))


def build():
    z = zipfile.ZipFile(FC.load()[DID]["path"])
    parts = {n: z.read(n) for n in z.namelist()}
    d = parts["word/document.xml"].decode("utf-8")
    i = d.find(NEEDLE)
    s = max(d.rfind("<w:p ", 0, i), d.rfind("<w:p>", 0, i)); e = d.find("</w:p>", i) + 6
    p = d[s:e]
    j = d.find("<w:sectPr", i)
    sect = re.search(r"<w:sectPr[ >].*?</w:sectPr>", d[j:], re.S).group(0)
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    paras = []
    for ind in range(LO, HI + 1, STEP):
        q = p
        if "<w:pPr>" in q:
            q = re.sub(r"<w:ind [^>]*/>", "", q, count=1)
            q = q.replace("<w:pPr>", '<w:pPr><w:ind w:right="%d"/>' % ind, 1)
        else:
            q = re.sub(r"(<w:p[ >][^>]*>|<w:p>)", lambda m: m.group(1) + '<w:pPr><w:ind w:right="%d"/></w:pPr>' % ind, q, count=1)
        paras.append(q + '<w:p><w:r><w:t>IND%d</w:t></w:r></w:p>' % ind)
    b0 = re.search(r"<w:body[^>]*>", d).end()
    parts["word/document.xml"] = (d[:b0] + "".join(paras) + sect + "</w:body></w:document>").encode("utf-8")
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "us_%s.docx" % DID.replace("/", "__"))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    return path


def verdicts(lines):
    out, cur = {}, []
    for t in lines:
        t = t.replace(" ", "").replace("　", "")
        if t.startswith("IND") and t[3:].lstrip("-").isdigit():
            joined = cur
            hit = next((l for l in joined if CONTEXT in l), None)
            out[int(t[3:])] = None if hit is None else hit.endswith(CONTEXT + X) or (CONTEXT + X) in hit
            cur = []
        else:
            cur.append(t)
    return out


def word_lines(pdf):
    import pymupdf
    res = []
    for pg in pymupdf.open(pdf):
        rows = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        rows.setdefault(round(c["origin"][1]), []).append((c["origin"][0], c["c"]))
        for k in sorted(rows):
            res.append("".join(c for _, c in sorted(rows[k])))
    return res


def oxi_lines(exe, docx):
    import subprocess, json, tempfile
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    res = []
    for p in json.load(open(g, encoding="utf-8"))["pages"]:
        rows = {}
        for c in p["glyphs"]:
            rows.setdefault(round(c["baseline"]), []).append(c)
        for k in sorted(rows):
            res.append("".join(c["char"] for c in sorted(rows[k], key=lambda c: c["x"])))
    return res


if __name__ == "__main__":
    import win32com.client
    p = build(); pdf = p[:-5] + ".pdf"
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    wv = verdicts(word_lines(pdf))
    ov = verdicts(oxi_lines(os.environ["PB_OXI"], p)) if os.environ.get("PB_OXI") else {}
    for ind in sorted(wv):
        print("ind %5d  Word %-5s Oxi %-5s %s" % (ind, wv[ind], ov.get(ind), "" if wv[ind] == ov.get(ind) else "<<"), flush=True)
