# -*- coding: utf-8 -*-
"""The footer paragraph after a table, on more documents than 0061c884.

For each document below, every footer part whose table is followed by a
paragraph gets that paragraph's space before set to 0 or 12pt (direct
w:spacing w:before). Read: page 1's footer-zone text baselines (below 80% of
the page) and the footer table's text, Word against Oxi; a correct model moves
everything above the trailing paragraph up by exactly the before.

    python tools/metrics/_pb_ftr_after_table_multi.py <exe>[;<exe>]
"""
import json, os, re, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_ftr_after_table_multi"
DOCS = ["en/administrative/0067cba813a366db", "en/administrative/001eef607583445c", "en/forms/004ec51fa3692fff", "en/technical/00a54bff7e15af78"]


def variant(doc, before):
    src = REPO / "pipeline_data/docx_corpus" / f"{doc}.docx"
    z = zipfile.ZipFile(src)
    parts = {n: z.read(n) for n in z.namelist()}
    for n in list(parts):
        if not re.match(r"word/footer\d*\.xml$", n):
            continue
        x = parts[n].decode("utf-8")
        m = re.search(r"</w:tbl>\s*<w:p([ >])", x)
        if not m:
            continue
        i = m.end()
        rest = x[i - 1:]
        # the paragraph's pPr: add (or replace) a spacing before
        j = x.find(">", i - 1) + 1 if rest.startswith(" ") else i
        para_end = x.find("</w:p>", j)
        body = x[j:para_end]
        sp = f'<w:spacing w:before="{before * 20}"/>'
        if "<w:pPr>" in body:
            body = re.sub(r"<w:spacing [^>]*/>", "", body, count=1)
            k = body.find("<w:pPr>") + len("<w:pPr>")
            ps = re.match(r"<w:pStyle [^>]*/>", body[k:])
            k += ps.end() if ps else 0
            body = body[:k] + sp + body[k:]
        else:
            body = f"<w:pPr>{sp}</w:pPr>" + body
        parts[n] = (x[:j] + body + x[para_end:]).encode("utf-8")
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{doc.replace('/', '_')}_b{before}.docx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    return p


def word(p):
    pdf = p.with_suffix(".pdf")
    if not pdf.exists():
        import win32com.client
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(str(p), ReadOnly=True); d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
        finally:
            w.Quit()
    import pymupdf
    pg = pymupdf.open(pdf)[0]
    h = pg.rect.height
    return sorted({round(c["origin"][1], 2) for b in pg.get_text("rawdict")["blocks"] for l in b.get("lines", [])
                   for s in l["spans"] for c in s["chars"] if c["origin"][1] > 0.8 * h and c["c"].strip()})


def oxi(exe, p):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, str(p), os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    pg = json.load(open(g, encoding="utf-8"))["pages"][0]
    return sorted({round(c["baseline"], 2) for c in pg["glyphs"] if c["baseline"] > 0.8 * pg["height"] and c["char"].strip()})


if __name__ == "__main__":
    exes = [e for e in sys.argv[1].split(";") if e]
    for doc in DOCS:
        for before in (0, 12):
            p = variant(doc, before)
            print(f"{doc.split('/')[-1][:8]} b{before:<2} W {word(p)}", flush=True)
            for e in exes:
                print(f"{'':12} O {oxi(e, p)}", flush=True)
