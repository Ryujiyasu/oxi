# -*- coding: utf-8 -*-
"""What is Word's ACTUAL line advance for Garamond 10pt, and Oxi's?

forms__002fbe2c6e5f24b5 (0.9949, one slip) accumulates ~1.7pt against Word down
p2 and drops its last line. Reading pitches off the real document's PDF was a
TRAP: bbox-top pitch varies with which glyphs each line holds (11.28 / 11.59 /
10.85 for the same nominal advance), because the bbox top is the TALLEST GLYPH,
not the line box. Meanwhile Oxi's own [LH] dump shows THREE candidate heights
(base=11.00, run_base=10.50, hhea=11.25) and the winner is not readable from
the outside.

So: arms of IDENTICAL repeated lines (same text -> same bbox geometry -> the
bbox-top pitch IS the advance, exactly), swept over the things the document
mixes:

    font     Garamond / Arial / Times New Roman, sz 10
    spacing  line=240 auto / line=276 auto (the docDefaults 1.15) / inherit
    gap      contextual pair (para after=0 -> the para-to-para step)

    python _pb_garapitch_gen.py gen
    python _pb_garapitch_gen.py pdf   # Word truth
    python _pb_garapitch_gen.py oxi   # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_garapitch")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

PGW, PGH, MARG = 12240, 15840, 1440
FONTS = ["Garamond", "Arial", "Times New Roman"]
SPACINGS = [("a240", 240), ("a276", 276)]
NLINES = 8
# One WRAPPED paragraph per arm too: within-paragraph advance can differ from
# the paragraph-to-paragraph step (spacing after / contextual gaps ride the
# latter). The wrap text repeats one word so every wrapped line is identical.
WRAP_WORDS = 60


def rpr(font):
    return ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
            '<w:sz w:val="20"/></w:rPr>' % (font, font))


def para(text, font, line, pbb=False):
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="%d"'
            ' w:lineRule="auto"/><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
            '<w:sz w:val="20"/></w:rPr></w:pPr><w:r>%s<w:t xml:space="preserve">%s'
            "</w:t></w:r></w:p>"
            % ("<w:pageBreakBefore/>" if pbb else "", line, font, font, rpr(font), text))


def arms():
    return [(f, sk, sv) for f in FONTS for sk, sv in SPACINGS]


def docx():
    return os.path.join(OUT, "garapitch.docx")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (font, _sk, sv) in enumerate(arms()):
        body.append(para("M%02d" % ai, font, sv, pbb=ai > 0))
        # (a) N single-line paragraphs, identical text: para-to-para step
        for j in range(NLINES):
            body.append(para("a%dP%d Hxg pqj kern" % (ai, j), font, sv))
        # (b) one wrapped paragraph, identical word: within-para line advance
        body.append(para(" ".join("a%dWx" % ai for _ in range(WRAP_WORDS)), font, sv))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="%d" w:h="%d"/>'
           '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
           % (PGW, PGH, MARG, MARG, MARG, MARG))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
              '<w:sz w:val="20"/></w:rPr></w:rPrDefault>'
              # the real document's docDefaults: after=200 line=276 auto. Arms
              # override both explicitly, so this only exercises inheritance.
              '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276"'
              ' w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style></w:styles>')
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(arms()), "arms")


def report(per, who):
    print("== %s ==" % who)
    print("%-16s %-6s %-10s %-10s" % ("font", "line", "para_step", "wrap_step"))
    for ai, (font, sk, _sv) in enumerate(arms()):
        g = per.get(ai)
        if not g:
            print("%-16s %-6s MISSING" % (font[:16], sk))
            continue
        pstep, wstep = g
        print("%-16s %-6s %-10s %-10s"
              % (font[:16], sk,
                 "%.3f" % pstep if pstep else "-",
                 "%.3f" % wstep if wstep else "-"))


def _steps(ys):
    """Median gap between successive identical lines."""
    ys = sorted(set(ys))
    if len(ys) < 2:
        return None
    gaps = [b - a for a, b in zip(ys, ys[1:])]
    gaps.sort()
    return gaps[len(gaps) // 2]


def pdf():
    import fitz
    import win32com.client as w
    out = docx().replace(".docx", ".pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx(), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(out, 17)
    finally:
        d.Close(False)
        app.Quit()
    doc = fitz.open(out)
    per = {}
    for ai, _a in enumerate(arms()):
        pys, wys = [], []
        for pi in range(doc.page_count):
            for bl in doc[pi].get_text("dict")["blocks"]:
                if bl["type"] != 0:
                    continue
                for ln in bl["lines"]:
                    t = "".join(s["text"] for s in ln["spans"]).strip()
                    if t.startswith("a%dP" % ai):
                        pys.append(ln["bbox"][1])
                    elif t.startswith("a%dWx" % ai):
                        wys.append(ln["bbox"][1])
        per[ai] = (_steps(pys), _steps(wys))
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "garapitch_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "gp"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    per = {}
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    for ai, _a in enumerate(arms()):
        pys, wys = [], []
        for pg in pages:
            for e in pg["elements"]:
                if e.get("type") != "text":
                    continue
                t = (e.get("text") or "").strip()
                if t.startswith("a%dP" % ai):
                    pys.append(e["y"])
                elif t.startswith("a%dWx" % ai):
                    wys.append(e["y"])
        per[ai] = (_steps(pys), _steps(wys))
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
