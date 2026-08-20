# -*- coding: utf-8 -*-
"""Does a Latin line holding U+00D7 (×) / U+00F7 (÷) keep its Latin line height?

creative__0158c02ae543d567 (44 pagination slips): every ÷/× line renders
h=18.00 in Oxi where the surrounding ArialMT 12pt ×1.15 body is 15.87 —
`kinsoku::is_cjk` claims 0x00D7/0x00F7 unconditionally ("Latin-1 math symbols
Word renders with East Asian font"), so the run routes down the eastAsia chain
→ S634 MS Mincho → 83/64 × 12 × 1.15 = 17.90. Word's own PDF draws those
lines in ArialMT at the plain 15.84 pitch (cre0158.pdf p9+).

Arms: font {Arial, Calibri, Times New Roman} × char {×, ÷, ½, plain}
× spacing {240, 276}, NLINES identical lines each so the bbox-top step IS the
advance (the _pb_garapitch method). Expectation if D7/F7 are the ambiguous
class (S1115 sibling): every arm = its plain control.

    python _pb_divmul_gen.py gen
    python _pb_divmul_gen.py pdf   # Word truth
    python _pb_divmul_gen.py oxi [ENV=..,..]
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_divmul")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

PGW, PGH, MARG = 12240, 15840, 1440
FONTS = ["Arial", "Calibri", "Times New Roman"]
CHARS = [("plain", ""), ("mul", "×"), ("div", "÷"), ("half", "½")]
SPACINGS = [("a240", 240), ("a276", 276)]
NLINES = 12
# The real document's eastAsia context: docDefaults eastAsiaTheme="minorHAnsi"
# (a LATIN face in the eastAsia slot). Without any eastAsia declaration the
# chain resolves to None and nothing inflates — the first probe cut proved
# that. An explicit Latin eastAsia reproduces the S634 substitution path
# (ukframework precedent: eastAsia="Times New Roman" → MS Mincho).
EASTASIA = "Calibri"


def arms():
    return [(f, ck, cc, sk, sv)
            for f in FONTS for ck, cc in CHARS for sk, sv in SPACINGS]


def docx():
    return os.path.join(OUT, "divmul.docx")


def para(ai, j, font, sym, line, pbb=False):
    mid = ("2 %s 1 " % sym) if sym else "2 v 1 "
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="%d"'
            ' w:lineRule="auto"/><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
            '<w:sz w:val="24"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/><w:sz w:val="24"/></w:rPr>'
            '<w:t xml:space="preserve">a%dP%d Hxg %spqj kern</w:t></w:r></w:p>'
            % ("<w:pageBreakBefore/>" if pbb else "", line, font, font,
               font, font, ai, j, mid))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (font, _ck, cc, _sk, sv) in enumerate(arms()):
        for j in range(NLINES):
            body.append(para(ai, j, font, cc, sv, pbb=(j == 0 and ai > 0)))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="%d" w:h="%d"/>'
           '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
           % (PGW, PGH, MARG, MARG, MARG, MARG))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:eastAsia="EAFONT"'
              ' w:hAnsi="Times New Roman"/>'
              '<w:sz w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US"/>'
              "</w:rPr></w:rPrDefault>"
              '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240"'
              ' w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style></w:styles>').replace("EAFONT", EASTASIA)
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(arms()), "arms")


def _steps(ys):
    ys = sorted(set(ys))
    if len(ys) < 2:
        return None
    gaps = [b - a for a, b in zip(ys, ys[1:])]
    return sum(gaps) / len(gaps)


def report(per, who):
    print("== %s ==" % who)
    print("%-16s %-6s %-5s %-9s %s" % ("font", "char", "line", "step", "faces"))
    for ai, (font, ck, _cc, sk, _sv) in enumerate(arms()):
        g = per.get(ai)
        if not g:
            print("%-16s %-6s %-5s MISSING" % (font[:16], ck, sk))
            continue
        step, faces = g
        print("%-16s %-6s %-5s %-9s %s"
              % (font[:16], ck, sk, "%.3f" % step if step else "-", faces))


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
        ys, faces = [], set()
        for pi in range(doc.page_count):
            for bl in doc[pi].get_text("dict")["blocks"]:
                if bl["type"] != 0:
                    continue
                for ln in bl["lines"]:
                    t = "".join(s["text"] for s in ln["spans"])
                    if t.startswith("a%dP" % ai):
                        ys.append(ln["bbox"][1])
                        faces |= set(s["font"] for s in ln["spans"])
        per[ai] = (_steps(ys), sorted(faces))
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "divmul_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "dm"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    per = {}
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    for ai, _a in enumerate(arms()):
        ys = []
        for pg in pages:
            for e in pg["elements"]:
                if e.get("type") != "text":
                    continue
                t = (e.get("text") or "").strip()
                if t.startswith("a%dP" % ai):
                    ys.append(e["y"])
        per[ai] = (_steps(ys), [])
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
