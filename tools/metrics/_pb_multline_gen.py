# -*- coding: utf-8 -*-
"""With a MULTIPLIED line rule, may the last line's leading overrun the page?

educational__00158a7d549f9f51 (0.8427) is a double-spaced thesis. Its within-page
paragraph pitches match Word to under 1pt, yet from p42 Oxi fits fewer lines per
page and the loss compounds in exact multiples of the 27.6pt pitch (+82.9 ->
+111 -> +147 -> +165 -> +192 -> +221). On p42 Word's LAST line box would end at
699.68 + 27.6 = 727.28, past the 720 text bottom, while its glyphs stop near 713.

So the question this sweeps: for line spacing 240 / 360 / 480 (single / 1.5 /
double), where does Word stop? Each arm is a run of one-line paragraphs walking
into the page bottom, so the readout is simply the last line Word kept.

  last_y + pitch  <= 720 for every spacing  -> Word reserves the whole line box,
                                               and the 00158a7d cause is elsewhere
                                               (the keepNext hypothesis)
  last_y + pitch  >  720 as the multiplier grows -> the trailing leading is
                                               allowed to spill, and Oxi's
                                               whole-box reservation is the bug

`ink` is what the glyph box actually occupies (PDF line bbox height), which is
the quantity that stays inside the margin if the spill rule is real.

NOT to be confused with `_pb_lastline_gen.py`, which asks the same-shaped
question about a TYPED GRID's last line (the S1152 phase x slack sweep). This
one is about the line-spacing MULTIPLIER on a no-grid Latin page.

    python _pb_multline_gen.py gen
    python _pb_multline_gen.py pdf   # Word truth
    python _pb_multline_gen.py oxi   # Oxi
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_multline")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

PGW, PGH, MARG = 12240, 15840, 1440   # Letter, 1in margins -> text band 72..720
TOP = MARG / 20.0
BOT = (PGH - MARG) / 20.0
# 240 = single, 360 = 1.5, 480 = double (w:line with lineRule="auto").
SPACINGS = [240, 360, 480]
NLINES = 60
# ★The first cut could not discriminate: in all three arms the box test and the
# ink test EXCLUDE the same next line, because the run's line positions are
# quantised by the pitch itself. The window where they disagree is only
# (pitch - ink) = 27.6 - 13.3 = 14.3pt wide for double spacing, and the real
# document's p42 line sits inside it (box bottom 727.28 > 720, ink ~713 <= 720).
# So shift the whole run in sub-pitch steps with a leading EXACT-height spacer
# and re-ask. PHASES are in points; 0 reproduces the first cut.
PHASES = [0, 4, 8, 12, 16, 20, 24]


def rpr():
    return ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="24"/></w:rPr>')


def spacer(pt):
    """A leading paragraph of EXACT height, to move the run off the pitch grid."""
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d"'
            ' w:lineRule="exact"/></w:pPr><w:r>%s<w:t>.</w:t></w:r></w:p>'
            % (int(pt * 20), rpr()))


def arms():
    return [(sp, ph) for sp in SPACINGS for ph in PHASES]


def para(text, line, pbb=False):
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="%d"'
            ' w:lineRule="auto"/></w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t>'
            "</w:r></w:p>" % ("<w:pageBreakBefore/>" if pbb else "", line, rpr(), text))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (sp, ph) in enumerate(arms()):
        body.append(para("M%02d" % ai, sp, pbb=ai > 0))
        if ph:
            body.append(spacer(ph))
        for j in range(NLINES):
            body.append(para("a%dL%02d" % (ai, j), sp))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="%d" w:h="%d"/>'
           '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
           % (PGW, PGH, MARG, MARG, MARG, MARG))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
              '<w:sz w:val="24"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
              ' w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style></w:styles>')
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(arms()), "arms x", NLINES, "lines")


def docx():
    return os.path.join(OUT, "multline.docx")


def report(per, who):
    print("== %s ==  (text band %.1f .. %.1f)" % (who, TOP, BOT))
    print("%-5s %-4s %-7s %-5s %-9s %-11s %-11s %s"
          % ("line", "ph", "pitch", "kept", "last_y", "box_bottom", "ink_bottom", "verdict"))
    for ai, (sp, ph) in enumerate(arms()):
        g = per.get(ai)
        if not g:
            print("%-5d %-4d MISSING" % (sp, ph))
            continue
        kept, _first_y, last_y, pitch, ink = g
        box = last_y + pitch
        ib = (last_y + ink) if ink else None
        # The discriminating case: the kept line's BOX overruns the text bottom
        # while its INK does not. Only an ink-based fit test can produce it.
        if box > BOT + 0.5:
            v = "INK-FIT (box overruns %+.2f)" % (box - BOT)
        elif ib is not None and ib > BOT + 0.5:
            v = "?? ink overruns"
        else:
            v = "box fits - undecided"
        print("%-5d %-4d %-7.2f %-5d %-9.2f %-11.2f %-11s %s"
              % (sp, ph, pitch, kept, last_y, box,
                 ("%.2f" % ib) if ib else "-", v))


def _collect(page_of, inks):
    per = {}
    for ai, _a in enumerate(arms()):
        st = page_of.get("M%02d" % ai)
        if not st:
            continue
        p0 = st[0]
        ys = [page_of[k][1] for k in (("a%dL%02d" % (ai, j)) for j in range(NLINES))
              if k in page_of and page_of[k][0] == p0]
        if len(ys) < 2:
            continue
        ys.sort()
        per[ai] = (len(ys), ys[0], ys[-1], ys[1] - ys[0], inks.get(ai))
    return per


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
    page_of, inks = {}, {}
    for pi in range(doc.page_count):
        for bl in doc[pi].get_text("dict")["blocks"]:
            if bl["type"] != 0:
                continue
            for ln in bl["lines"]:
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if not t:
                    continue
                page_of.setdefault(t, (pi, round(ln["bbox"][1], 2)))
                for ai in range(len(arms())):
                    if t.startswith("a%dL" % ai):
                        inks.setdefault(ai, round(ln["bbox"][3] - ln["bbox"][1], 2))
    report(_collect(page_of, inks), "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "multline_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "ml"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    page_of = {}
    for pi, pg in enumerate(json.load(open(out, encoding="utf-8"))["pages"]):
        for e in pg["elements"]:
            if e.get("type") != "text":
                continue
            t = (e.get("text") or "").strip()
            if t:
                page_of.setdefault(t, (pi, round(e["y"], 2)))
    report(_collect(page_of, {}), "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
