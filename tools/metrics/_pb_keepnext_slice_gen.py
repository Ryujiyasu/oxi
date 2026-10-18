# -*- coding: utf-8 -*-
"""Faithful multi-paragraph slice for a keepNext page-end decision.

Keeps a run of real paragraphs (from the one containing START to the one
containing END, inclusive) with the document's styles / settings / section, and
puts an exact-height spacer paragraph in front of them; one document per spacer
height, so the run slides down to the page end.  Reports, per height, the page
of each probe word.  For technical__0056b52f: does Word keep «event counter» (a
run-in Heading3 entry, keepNext) on the page when «(Added 1993)» after it moves?

    python tools/metrics/_pb_keepnext_slice_gen.py <doc_id> "<START>" "<END>" "<word1>|<word2>|..." <h0,h1,step>
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import feature_census as FC  # noqa: E402

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_keepnext_slice")


def para_bounds(d, i):
    s = max(d.rfind("<w:p ", 0, i), d.rfind("<w:p>", 0, i))
    e = d.find("</w:p>", i) + len("</w:p>")
    return s, e


def build(did, start, end, height_tw):
    z = zipfile.ZipFile(FC.load()[did]["path"])
    parts = {n: z.read(n) for n in z.namelist()}
    d = parts["word/document.xml"].decode("utf-8")
    i0 = d.find(start)
    s0, _ = para_bounds(d, i0)
    i1 = d.find(end, i0)
    _, e1 = para_bounds(d, i1)
    run = d[s0:e1]
    sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", d, re.S)[-1]
    spacer = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="exact"/></w:pPr>'
              '<w:r><w:t>SPACER</w:t></w:r></w:p>' % height_tw)
    b0 = re.search(r"<w:body[^>]*>", d).end()
    parts["word/document.xml"] = (d[:b0] + spacer + run + sect + "</w:body></w:document>").encode("utf-8")
    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, "%s_h%d.docx" % (did.replace("/", "__"), height_tw))
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    return p


def pages_of(docx, words):
    import fitz, win32com.client
    pdf = docx[:-5] + ".pdf"
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    res = {}
    for i, pg in enumerate(fitz.open(pdf)):
        t = pg.get_text().replace("\n", " ")
        for wd in words:
            if wd not in res and wd in t:
                res[wd] = i + 1
    return res


def main():
    did, start, end, words = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4].split("|")
    h0, h1, step = map(int, sys.argv[5].split(","))
    prev = None
    for h in range(h0, h1, step):
        r = pages_of(build(did, start, end, h), words)
        key = tuple(r.get(w) for w in words)
        if key != prev:
            print("spacer %5dtw (%.1fpt): %s" % (h, h / 20, "  ".join("%s=p%s" % (w, r.get(w)) for w in words)), flush=True)
        prev = key


if __name__ == "__main__":
    main()
