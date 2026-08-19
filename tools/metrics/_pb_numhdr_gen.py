# -*- coding: utf-8 -*-
"""The forms numbered-heading step: where does Oxi's +0.5 come from?

forms__002fbe2c p2: the two spans [empty, empty, numbered heading] measure
33.73/33.84 in Word (= 3 x plain Garamond 11.25, the whitespace-tab law) but
34.25 in Oxi. Everything plausible has been ruled out by measurement: the
marker resolves to Garamond Bold 10 (correct, [MARKER] trace), Garamond Bold's
metrics equal the regular's (font file and table), and line_height() already
reports hhea=11.25 for the heading line. So the +0.5 is added by some growth
path AFTER line height -- this repro isolates the shape so the flag bisect can
name it.

Shape: plain Garamond 10pt paragraphs, two empties, then a numbered paragraph
(numPr, lvl rPr = hint-default + bold only, runs Garamond bold), then plain
again -- the forms structure exactly.

    python _pb_numhdr_gen.py gen
    python _pb_numhdr_gen.py pdf                 # Word truth
    python _pb_numhdr_gen.py oxi                 # Oxi default
    python _pb_numhdr_gen.py oxi OXI_S1112_DISABLE=1   # flag bisect arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_numhdr")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

PGW, PGH, MARG = 12240, 15840, 1440
NREP = 8   # repeat the [plain, empty, empty, numbered] block


def rpr(bold=False):
    return ('<w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond"/>%s'
            '<w:sz w:val="20"/></w:rPr>' % ("<w:b/>" if bold else ""))


def para(text, bold=False, num=False):
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/>%s</w:pPr>%s</w:p>'
            % ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' if num else "",
               rpr(bold),
               ('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr(bold), text))
               if text else ""))


def docx():
    return os.path.join(OUT, "numhdr.docx")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    # rotate the H variant: full (bold+num) / bold only / num only / plain
    variants = [(True, True), (True, False), (False, True), (False, False)]
    for j in range(NREP):
        b, n = variants[j % len(variants)]
        body.append(para("k%dA plain body line" % j))
        body.append(para(""))
        body.append(para(""))
        body.append(para("k%dH heading" % j, bold=b, num=n))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="%d" w:h="%d"/>'
           '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
           % (PGW, PGH, MARG, MARG, MARG, MARG))
    numbering = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + NS + ">"
                 '<w:abstractNum w:abstractNumId="0">'
                 '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
                 '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>'
                 '<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>'
                 '<w:rPr><w:rFonts w:hint="default"/><w:b/><w:color w:val="auto"/></w:rPr>'
                 "</w:lvl></w:abstractNum>"
                 '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              # forms-faithful docDefaults: Calibri 11 (the theme minor font)
              '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
              '<w:sz w:val="22"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240"'
              ' w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style></w:styles>')
    ct = CT.replace(
        "</Types>",
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/'
             '2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/'
             '2006/relationships/numbering" Target="numbering.xml"/></Relationships>')
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/numbering.xml", numbering)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), NREP, "blocks")


def report(pos, who):
    print("== %s ==  (span = A(k) -> H(k): two empties + the numbered heading)" % who)
    spans = []
    for j in range(NREP):
        a = pos.get("k%dA" % j)
        h = pos.get("k%dH" % j)
        if a and h and a[0] == h[0]:
            spans.append(h[1] - a[1])
    for j, sp in enumerate(spans):
        print("  block %d span %.3f" % (j, sp))
    if spans:
        spans.sort()
        print("  median %.3f   (plain law predicts 3 x 11.25 = 33.75)"
              % spans[len(spans) // 2])


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
    pos = {}
    for pi in range(doc.page_count):
        for bl in doc[pi].get_text("dict")["blocks"]:
            if bl["type"] != 0:
                continue
            for ln in bl["lines"]:
                t = "".join(s["text"] for s in ln["spans"]).strip()
                for j in range(NREP):
                    for tag in ("k%dA" % j, "k%dH" % j):
                        if tag in t:
                            pos.setdefault(tag, (pi, round(ln["bbox"][1], 2)))
    report(pos, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "numhdr_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "nh"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pos = {}
    for pi, pg in enumerate(json.load(open(out, encoding="utf-8"))["pages"]):
        for e in pg["elements"]:
            t = (e.get("text") or "")
            for j in range(NREP):
                for tag in ("k%dA" % j, "k%dH" % j):
                    if tag in t:
                        pos.setdefault(tag, (pi, round(e["y"], 2)))
    report(pos, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
