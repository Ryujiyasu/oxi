# -*- coding: utf-8 -*-
"""How much does a keepNext heading demand before it will stay on the page?

educational__00158a7d549f9f51 p42: Oxi stops with 86.04pt of the page free and
moves a `Heading` (keepNext + keepLines + spacing before=200) plus its whole
following paragraph to the next page. Word puts the heading AND TWO body lines
there, its last line landing ink-bottom 712.96 inside the 720 text bottom.

`_pb_lastline_gen.py` already showed the plain line rule -- a line fits when
`top + UNMULTIPLIED line height <= text bottom`, so the multiplier's trailing
leading may overrun -- and that Oxi matches Word on it in 21/21 arms. So the
remaining suspect is the keepNext GROUP: what does Word require to be able to
leave the heading behind, and how many following lines come with it?

Each arm walks the heading down the page: F double-spaced filler lines, then an
optional sub-pitch phase spacer, then the heading, then a long paragraph. The
readout is whether the heading stayed and how many body lines came with it.

  heading stays with k>=1 body lines whenever heading+1 line fits by the ink
  rule                       -> Word needs exactly ONE line, and Oxi is
                                over-reserving (full boxes for the group)
  heading only moves when the group cannot fit 2 lines
                             -> the requirement is two (a widow rule)

    python _pb_keepnext_gen.py gen
    python _pb_keepnext_gen.py pdf   # Word truth
    python _pb_keepnext_gen.py oxi   # Oxi
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_keepnext")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

PGW, PGH, MARG = 12240, 15840, 1440
TOP = MARG / 20.0
BOT = (PGH - MARG) / 20.0
LINE = 480              # double, as in the thesis
PITCH = 27.6
BEFORE = 200            # twips, the Heading2/3/4 spacing before = 10pt
FILLERS = [18, 19, 20, 21, 22, 23]
PHASES = [0, 7, 14, 21]
BODYLINES = 8


def rpr():
    return ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="24"/></w:rPr>')


def para(text, pbb=False):
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="%d"'
            ' w:lineRule="auto"/></w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t>'
            "</w:r></w:p>" % ("<w:pageBreakBefore/>" if pbb else "", LINE, rpr(), text))


def spacer(pt):
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d"'
            ' w:lineRule="exact"/></w:pPr><w:r>%s<w:t>.</w:t></w:r></w:p>'
            % (int(pt * 20), rpr()))


def heading(text):
    """Heading2/3/4 shape: keepNext + keepLines + spacing before, same line rule."""
    return ('<w:p><w:pPr><w:keepNext/><w:keepLines/>'
            '<w:spacing w:before="%d" w:after="0" w:line="%d" w:lineRule="auto"/>'
            "</w:pPr><w:r>%s<w:t>%s</w:t></w:r></w:p>" % (BEFORE, LINE, rpr(), text))


def body(ai):
    """A long paragraph; each line is tagged so the kept count is readable."""
    words = " ".join("a%dB%02d" % (ai, j) for j in range(BODYLINES * 6))
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d"'
            ' w:lineRule="auto"/></w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t>'
            "</w:r></w:p>" % (LINE, rpr(), words))


def arms():
    return [(f, p) for f in FILLERS for p in PHASES]


def docx():
    return os.path.join(OUT, "keepnext.docx")


def gen():
    os.makedirs(OUT, exist_ok=True)
    b = []
    for ai, (f, ph) in enumerate(arms()):
        b.append(para("M%02d" % ai, pbb=ai > 0))
        for j in range(f):
            b.append(para("a%dF%02d" % (ai, j)))
        if ph:
            b.append(spacer(ph))
        b.append(heading("H%02d heading" % ai))
        b.append(body(ai))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(b) +
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
    print("wrote", docx(), len(arms()), "arms")


def report(per, who):
    print("== %s ==  (text band %.1f .. %.1f, pitch %.1f, before %.1f)"
          % (who, TOP, BOT, PITCH, BEFORE / 20.0))
    print("%-4s %-4s %-10s %-9s %-8s %s"
          % ("fill", "ph", "last_filler", "free", "heading", "body lines kept"))
    for ai, (f, ph) in enumerate(arms()):
        g = per.get(ai)
        if not g:
            print("%-4d %-4d MISSING" % (f, ph))
            continue
        lastf, hstay, nbody, hy = g
        # free space below the last filler's line box
        free = BOT - (lastf + PITCH) if lastf else float("nan")
        print("%-4d %-4d %-10.2f %-9.2f %-8s %s"
              % (f, ph, lastf, free,
                 ("stay %.1f" % hy) if hstay else "MOVED",
                 nbody if hstay else "-"))


def _collect(page_of):
    per = {}
    for ai, (f, _ph) in enumerate(arms()):
        st = page_of.get("M%02d" % ai)
        if not st:
            continue
        p0 = st[0]
        fk = [page_of[k][1] for k in (("a%dF%02d" % (ai, j)) for j in range(f))
              if k in page_of and page_of[k][0] == p0]
        h = page_of.get("H%02d" % ai)
        hstay = bool(h and h[0] == p0)
        ys = {page_of[k][1] for k in (("a%dB%02d" % (ai, j)) for j in range(BODYLINES * 6))
              if k in page_of and page_of[k][0] == p0}
        per[ai] = (max(fk) if fk else 0.0, hstay, len(ys), h[1] if h else 0.0)
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
    page_of = {}
    for pi in range(doc.page_count):
        for bl in doc[pi].get_text("dict")["blocks"]:
            if bl["type"] != 0:
                continue
            for ln in bl["lines"]:
                y = round(ln["bbox"][1], 2)
                for sp in ln["spans"]:
                    for tok in sp["text"].split():
                        page_of.setdefault(tok.strip(), (pi, y))
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if t.startswith("H") and "heading" in t:
                    page_of.setdefault(t.split()[0], (pi, y))
    report(_collect(page_of), "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "keepnext_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "kn"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    page_of = {}
    for pi, pg in enumerate(json.load(open(out, encoding="utf-8"))["pages"]):
        for e in pg["elements"]:
            if e.get("type") != "text":
                continue
            y = round(e["y"], 2)
            for tok in (e.get("text") or "").split():
                page_of.setdefault(tok.strip(), (pi, y))
    report(_collect(page_of), "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
