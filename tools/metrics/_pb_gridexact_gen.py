# -*- coding: utf-8 -*-
"""At a line height EXACTLY equal to the grid pitch, how many cells does Word use?

_pb_gridcell_gen.py bracketed Word's rule to "natural <= pitch -> one cell", but
every arm sat a little under or a little over; the exact-equality case is the one
that matters.  Oxi's snap is

    cells = floor(h / pitch + 1.0)

which agrees with ceil(h / pitch) everywhere EXCEPT at an exact multiple, where
it returns one cell too many.  That single case is what makes
educational__0214ac95 (UD Digi Kyokasho 12pt -> 12 * 1.15674 * 83/64 = 18.000pt
against an 18pt pitch) come out three pages where Word gives two, and it is the
reason the CJK half of the S1142 font sweep is held.

Hitting equality needs a height that is a whole number of twips: MS Mincho's
83/64 inflation makes 16pt exactly 20.75pt = 415 twips, so a section with
linePitch=415 puts natural == pitch on the nose.  The arms step the pitch one
twip either side of that to show the transition rather than a single point.

  python _pb_gridexact_gen.py gen
  python _pb_gridexact_gen.py pdf      # Word truth
  python _pb_gridexact_gen.py oxi      # Oxi, same arms
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_gridexact")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

# MS Mincho 16pt: 1.0em natural, Word draws CJK at 83/64 -> 20.75pt = 415 twips.
SIZE_HP = 32          # half-points
NATURAL = 20.75
# (arm, linePitch twips): 413 = pitch below natural, 415 = exactly natural,
# 417 = pitch above natural
ARMS = [("p413", 413), ("p414", 414), ("p415", 415), ("p416", 416), ("p417", 417)]
SENT_JA = "この文書は行グリッドの一マスに収まる高さの上限を測るための本文です。"


def docx():
    return os.path.join(OUT, "gridexact.docx")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (name, pitch) in enumerate(ARMS):
        # each arm is its own SECTION so it can carry its own linePitch
        body.append(
            '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"'
            ' w:hAnsi="Arial"/><w:sz w:val="14"/></w:rPr><w:t>A%02dZ</w:t>'
            "</w:r></w:p>" % ai)
        body.append(
            '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="MS Mincho" w:hAnsi="MS Mincho" w:eastAsia="MS Mincho"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (SIZE_HP, SIZE_HP, SENT_JA * 10))
        sect = ('<w:pgSz w:w="11907" w:h="16839"/>'
                '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
                'w:header="720" w:footer="720" w:gutter="0"/>'
                '<w:docGrid w:type="lines" w:linePitch="%d"/>' % pitch)
        if ai < len(ARMS) - 1:
            # a section break carries the previous section's sectPr in a paragraph
            body.append('<w:p><w:pPr><w:sectPr>' + sect + "</w:sectPr></w:pPr></w:p>")
        else:
            tail = "<w:sectPr>" + sect + "</w:sectPr>"
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) + tail + "</w:body></w:document>")
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:eastAsia="MS Mincho"'
              ' w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/><w:rPr><w:sz w:val="21"/></w:rPr></w:style>'
              "</w:styles>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms; natural =", NATURAL, "pt")


def report(per, who):
    print("== %s ==  (MS Mincho %.0fpt, natural %.2fpt)" % (who, SIZE_HP / 2.0, NATURAL))
    print("%-6s %9s %9s %8s %9s" % ("arm", "pitch pt", "lines", "gap", "cells"))
    for ai, (name, pitch) in enumerate(ARMS):
        ys = per.get(ai) or []
        if len(ys) < 3:
            print("%-6s MISSING (%d lines)" % (name, len(ys)))
            continue
        p = pitch / 20.0
        gap = (ys[-1] - ys[0]) / (len(ys) - 1)
        print("%-6s %9.2f %9d %8.3f %9.2f" % (name, p, len(ys), gap, gap / p))


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
    page_of = {}
    for pi in range(doc.page_count):
        for m in re.finditer(r"A(\d\d)Z", doc[pi].get_text()):
            page_of.setdefault(int(m.group(1)), pi)
    per = {}
    for ai in range(len(ARMS)):
        pi = page_of.get(ai)
        if pi is None:
            continue
        ys = set()
        for bl in doc[pi].get_text("dict")["blocks"]:
            for ln in bl.get("lines", []):
                for sp in ln["spans"]:
                    if abs(sp["size"] - SIZE_HP / 2.0) < 0.3 and sp["text"].strip():
                        ys.add(round(sp["origin"][1], 3))
                        break
        per[ai] = sorted(ys)
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "gridexact_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "ge"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    page_of = {}
    for pi, pg in enumerate(pages):
        for e in pg["elements"]:
            m = re.fullmatch(r"A(\d\d)Z", (e.get("text") or "").strip())
            if m:
                page_of.setdefault(int(m.group(1)), pi)
    per = {}
    for ai in range(len(ARMS)):
        pi = page_of.get(ai)
        if pi is None:
            continue
        per[ai] = sorted({round(e["y"], 3) for e in pages[pi]["elements"]
                          if e.get("type") == "text" and (e.get("text") or "").strip()
                          and abs((e.get("font_size") or 0) - SIZE_HP / 2.0) < 0.3})
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
