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
# MIXED arms (2026-08-20): forms__002fbe2c's whole p2 drift lives in two spans
# whose heading line mixes Garamond with Arial Bold. Ink-corrected, Word's
# advance for that line reads ~11.28 (= the Garamond-ish uniform pitch) while
# Oxi bumps it to 11.75 -- so: which fragment governs a mixed line's advance?
# Each mixed arm is a run of identical lines "gara-word ARIAL-BOLD-word
# gara-word", so the pitch is again exact. The empty arms interleave EMPTY
# paragraphs (the other half of those spans) between marked lines: the pitch
# then reads (marked k+1) - (marked k) = empty + marked advance.
# 2026-08-20 second cut: the x1.15 window separated the candidate models by
# only 0.037pt. x1.5 and x2.0 separate them by 0.25-0.5pt:
#   model A  1.2em x factor                    -> 18.0 / 24.0      (G+A)
#   model B  1.2·fs + (factor-1) x hhea_max    -> 17.7495 / 23.499 (G+A)
#   model C  factor x hhea_max, floor 1.2em    -> 17.2485 / 22.998 (G+A)
# and Cambria (hhea 11.724) vs Arial (11.499) pins WHICH hhea enters the term.
MIXED = [("Garamond", "Arial"), ("Garamond", "Cambria"),
         ("Garamond", "Garamond"), ("Garamond", "!Arial"),
         # 2026-08-20 fourth cut: in the REAL forms document the only foreign-
         # font fragment on the heading line is the marker-suffix TAB (a space,
         # Arial-BoldMT in Word's own PDF) -- and the ink-corrected span reads
         # 33.75-33.84 = the PLAIN Garamond height, not the composed 34.51.
         # Hypothesis: a WHITESPACE-ONLY fragment does not join the line-height
         # composition. "~Name" = the second run is a single space in Name.
         ("Garamond", "~Arial"), ("Garamond", "~Cambria")]
MIX_SPACINGS = [("a240", 240), ("a276", 276), ("a360", 360), ("a480", 480)]
NLINES = 20
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
    base = [("plain", f, None, sk, sv) for f in FONTS for sk, sv in SPACINGS]
    mixed = [("mixed", fa, fb, sk, sv) for fa, fb in MIXED for sk, sv in MIX_SPACINGS]
    empty = [("empty", f, None, "a240", 240) for f in FONTS]
    # size scaling: if the mixed advance is 1.2em it must be 14.4 at 12pt
    big = [("mixed12", "Garamond", "Arial", "a240", 240)]
    # 2026-08-20 third cut: Cambria breaks the flat 1.2em (G+C@10 = 12.120).
    # Sweep the size at x1.0 for the Cambria pair, add plain Cambria/Calibri
    # controls and the Calibri pair -- K_font's functional form falls out.
    sz = ([("mixsz", "Garamond", "Cambria", "s%d" % z, z) for z in (8, 10, 12, 14, 16)]
          + [("mixsz", "Garamond", "Calibri", "s%d" % z, z) for z in (10, 12)]
          + [("plainsz", "Cambria", None, "s%d" % z, z) for z in (10, 12)]
          + [("plainsz", "Calibri", None, "s%d" % z, z) for z in (10, 12)])
    return base + mixed + empty + big + sz


def docx():
    return os.path.join(OUT, "garapitch.docx")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (kind, font, fb, _sk, sv) in enumerate(arms()):
        body.append(para("M%02d" % ai, font, sv, pbb=ai > 0))
        if kind == "plain":
            for j in range(NLINES):
                body.append(para("a%dP%d Hxg pqj kern" % (ai, j), font, sv))
            body.append(para(" ".join("a%dWx" % ai for _ in range(WRAP_WORDS)), font, sv))
        elif kind in ("mixsz", "plainsz"):
            z = sv * 2  # half-points
            if kind == "plainsz":
                for j in range(NLINES):
                    body.append(
                        '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
                        ' w:lineRule="auto"/><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
                        '<w:sz w:val="%d"/></w:rPr></w:pPr>'
                        '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
                        '<w:sz w:val="%d"/></w:rPr><w:t xml:space="preserve">a%dP%d Hxg pqj</w:t></w:r></w:p>'
                        % (font, font, z, font, font, z, ai, j))
            else:
                for j in range(NLINES):
                    body.append(
                        '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
                        ' w:lineRule="auto"/><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
                        '<w:sz w:val="%d"/></w:rPr></w:pPr>'
                        '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
                        '<w:sz w:val="%d"/></w:rPr><w:t xml:space="preserve">a%dP%d Hxg </w:t></w:r>'
                        '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
                        '<w:sz w:val="%d"/></w:rPr><w:t xml:space="preserve">BOLD</w:t></w:r>'
                        '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
                        '<w:sz w:val="%d"/></w:rPr><w:t xml:space="preserve"> pqj</w:t></w:r></w:p>'
                        % (font, font, z, font, font, z, ai, j, fb, fb, z, font, font, z))
        elif kind == "mixed12":
            for j in range(NLINES):
                body.append(
                    '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
                    ' w:lineRule="auto"/><w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond"/>'
                    '<w:sz w:val="24"/></w:rPr></w:pPr>'
                    '<w:r><w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond"/>'
                    '<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">a%dP%d Hxg </w:t></w:r>'
                    '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/>'
                    '<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">BOLD</w:t></w:r>'
                    '<w:r><w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond"/>'
                    '<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve"> pqj kern</w:t></w:r></w:p>'
                    % (ai, j))
        elif kind == "mixed":
            # identical mixed lines: FONT text, one styled fb word, FONT tail.
            # fb "!Name" = REGULAR (no bold) second family.
            nobold = fb.startswith("!") or fb.startswith("~")
            ws_only = fb.startswith("~")
            fb2 = fb.lstrip("!~")
            for j in range(NLINES):
                body.append(
                    '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d"'
                    ' w:lineRule="auto"/><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
                    '<w:sz w:val="20"/></w:rPr></w:pPr>'
                    "<w:r>%s<w:t xml:space=\"preserve\">a%dP%d Hxg </w:t></w:r>"
                    '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>%s'
                    '<w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
                    "<w:r>%s<w:t xml:space=\"preserve\"> pqj kern</w:t></w:r></w:p>"
                    % (sv, font, font, rpr(font), ai, j, fb2, fb2,
                       "" if nobold else "<w:b/>", " " if ws_only else "BOLD",
                       rpr(font)))
        else:
            # marked line / EMPTY / marked line / EMPTY ... : the step between
            # marked lines = one marked + one empty advance
            for j in range(NLINES):
                body.append(para("a%dP%d Hxg pqj kern" % (ai, j), font, sv))
                body.append(para("", font, sv))
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
    print("%-7s %-16s %-16s %-6s %-10s %-10s"
          % ("kind", "font", "second", "line", "para_step", "wrap_step"))
    for ai, (kind, font, fb, sk, _sv) in enumerate(arms()):
        g = per.get(ai)
        if not g:
            print("%-7s %-16s %-16s %-6s MISSING" % (kind, font[:16], (fb or "")[:16], sk))
            continue
        pstep, wstep = g
        print("%-7s %-16s %-16s %-6s %-10s %-10s"
              % (kind, font[:16], (fb or "")[:16], sk,
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
