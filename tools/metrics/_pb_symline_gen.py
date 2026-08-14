# -*- coding: utf-8 -*-
"""Which codepoints make Word give a LATIN run a CJK-sized line?

reports__0020157f's checkbox rows are `<w:rFonts w:ascii="Arial" .../>` runs of
one 18pt symbol followed by 10pt Arial text.  Oxi sizes them:

    U+25A1 WHITE SQUARE       23.34 = 18 x 1.297  (the CJK win*83/64 box)
    U+221A SQUARE ROOT        20.70 = 18 x 1.1504 (Arial hhea)

while Word's own advances across the four rows are ~21.0 / 21.0 / 20.25, i.e.
Arial for both.  The 2.3pt error on each square repeats down a 65-border form
table and is what finally pushes that document's br-page stub over the content
bottom by 0.36pt (a blank page, the last PASS->FAIL of the 2026-08-13 bundle).

Rather than special-case one codepoint, this sweeps a range of symbol blocks in
a Latin run and reads Word's own line height for each, with a true CJK arm and a
plain Latin arm as the two controls.  Each arm is one page holding REPEAT copies
between markers, so the per-copy height comes out to +-0.05pt.

★The plain arms do NOT reproduce the specimen: with an explicit Arial in
docDefaults, Oxi gives U+25A1 the Arial line (21.000) like Word.  The specimen
instead resolves its fonts through the THEME
(`<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" .../>`)
whose `<a:ea typeface=""/>` is EMPTY.  The `theme` variant below mirrors exactly
that -- same docDefaults, the specimen's own theme1.xml part, runs still naming
Arial explicitly -- so the two documents differ in one thing only.

  python _pb_symline_gen.py gen  [theme]
  python _pb_symline_gen.py read [theme]      # Word COM truth
  python _pb_symline_gen.py oxi  ""  [theme]  # Oxi, same arms
"""
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_symline")
THEME = False
# ★GRID variant: the same arms under a <w:docGrid>. The plain/theme arms have no
# docGrid at all, which is the NO-GRID body path; a real corpus form usually has
# one, and the line height is computed by a DIFFERENT function there. Without
# this variant the grid wiring would ship unmeasured.
GRID = False
# ★CELL variant: the REPEAT subject paragraphs go inside a one-cell table, with
# the markers still in the body.  A cell line height has its own clamps and the
# row can pin it via trHeight, so whether Word applies the SAME fallback rule
# there is a separate question from the body arms — S1119 deliberately shipped
# without wiring cells until this variant answers it.  No trHeight is set, so
# the row is free to grow with its content.
CELL = False
SPECIMEN = os.path.join(REPO, "pipeline_data", "docx_corpus", "en", "reports",
                        "0020157f48ee08b2.docx")


def docx():
    return os.path.join(OUT, "symline%s%s%s.docx"
                        % ("_theme" if THEME else "", ("_grid" if GRID else "") + ("_cell" if CELL else ""),
                           "_cjk" if WITH_CJK else ""))


# docDefaults exactly as the specimen writes them, so the only difference from
# the plain arms is where the font names come from.
THEME_STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + "%s" + ">"
    "<w:docDefaults><w:rPrDefault><w:rPr>"
    '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"'
    ' w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/>'
    "</w:rPr></w:rPrDefault>"
    '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    "</w:pPrDefault></w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    "<w:name w:val=\"Normal\"/></w:style></w:styles>")
THEME_CT = (
    '<Override PartName="/word/theme/theme1.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>')
THEME_REL = (
    '<Relationship Id="rIdTh" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS, STYLES  # noqa: E402

REPEAT = 8
SYM_SZ, TXT_SZ = 36, 20                 # half-points: 18pt symbol, 10pt text

# (codepoint, name) -- the specimen's two, their neighbours, and both controls
CHARS = [
    (0x0041, "A LATIN control"),
    (0x00A7, "SECTION SIGN"),
    (0x2022, "BULLET"),
    (0x2190, "LEFTWARDS ARROW"),
    (0x221A, "SQUARE ROOT"),
    (0x2460, "CIRCLED ONE"),
    (0x25A0, "BLACK SQUARE"),
    (0x25A1, "WHITE SQUARE"),
    (0x25CB, "WHITE CIRCLE"),
    (0x25C6, "BLACK DIAMOND"),
    (0x2610, "BALLOT BOX"),
    (0x2713, "CHECK MARK"),
    # The two shapes the CORPUS actually carries (census 2026-08-14:
    # forms__002f81ab has 16x Cambria U+2605, forms__001ae487 has 4x
    # Calibri U+2610). Everything above was chosen from the specimen
    # `reports__0020157f`; these two are what the fix has to get right.
    (0x2605, "BLACK STAR"),
]
# ★These two must live in their OWN document.  A single real CJK character
# anywhere in the body sets Oxi's doc_body_has_real_cjk, which switches off
# every ambiguous-class carve-out (S801/S830/S888/S951/S966/S1103) at once — the
# first cut of this probe put them beside the symbols and measured a flat CJK
# line for all 14, which says nothing about the Latin-document path.
CJK_CONTROLS = [
    (0x3007, "IDEOGRAPHIC ZERO"),
    (0x4E00, "CJK control"),
]
FONTS = ["Arial", "Calibri", "Cambria"]
WITH_CJK = False


def chars():
    return CHARS + (CJK_CONTROLS if WITH_CJK else [])


def arms():
    """Each font opens with its OWN marker-only control arm (cp None).

    The first cut used a single Arial control for every font, which inflated
    every Calibri per-copy value by (Calibri marker - Arial marker)/REPEAT =
    0.12pt -- enough to blur the 0.375pt fallback steps this probe is after.
    """
    return [(f, cp, nm) for f in FONTS for cp, nm in [(None, "control")] + chars()]


def rpr(font, sz):
    return ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            % (font, font, font, sz, sz))


def ppr(font, pbb=False):
    return ('<w:pPr>%s<w:widowControl w:val="0"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '%s</w:pPr>' % ("<w:pageBreakBefore/>" if pbb else "", rpr(font, TXT_SZ)))


def subject(font, cp):
    return ('<w:p>%s<w:r>%s<w:t>&#x%X;</w:t></w:r>'
            '<w:r>%s<w:t xml:space="preserve">x</w:t></w:r></w:p>'
            % (ppr(font), rpr(font, SYM_SZ), cp, rpr(font, TXT_SZ)))


def marker(tag, font, pbb=False):
    return "<w:p>%s<w:r>%s<w:t>%s</w:t></w:r></w:p>" % (ppr(font, pbb), rpr(font, TXT_SZ), tag)


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (font, cp, _nm) in enumerate(arms()):
        body.append(marker("M%02dS" % ai, font, True))
        subs = [subject(font, cp) for _ in range(REPEAT if cp is not None else 0)]
        if CELL:
            # One cell holding every subject paragraph. The control arm (cp None)
            # still emits the table with a single empty paragraph so the cell's
            # own padding cancels in (span - control).
            inner = "".join(subs) if subs else ("<w:p>%s</w:p>" % ppr(font))
            body.append(
                '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
                '<w:tblLayout w:type="fixed"/></w:tblPr>'
                '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
                '<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>'
                + inner + "</w:tc></w:tr></w:tbl>")
        else:
            body.extend(subs)
        body.append(marker("M%02dE" % ai, font))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           + ('<w:docGrid w:linePitch="360"/>' if GRID else '')
           + '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>')
    ct, drels, styles = CT, DRELS, STYLES
    if THEME:
        ct = CT.replace("</Types>", THEME_CT + "</Types>")
        drels = DRELS.replace("</Relationships>", THEME_REL + "</Relationships>")
        styles = THEME_STYLES % NS
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/styles.xml", styles)
        if THEME:
            # the specimen's own theme part, so `<a:ea typeface=""/>` is verbatim
            z.writestr("word/theme/theme1.xml",
                       zipfile.ZipFile(SPECIMEN).read("word/theme/theme1.xml"))
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(arms()), "arms x", REPEAT, "theme" if THEME else "")


def report(spans, who):
    base = {}
    for ai, (font, cp, _nm) in enumerate(arms()):
        if cp is None and spans.get(ai) is not None:
            base[font] = spans[ai]
    print("%s   marker-only spans %s"
          % (who, {f: round(v, 2) for f, v in base.items()}))
    print("%-9s %-8s %-18s %9s" % ("font", "cp", "name", "per copy"))
    for ai, (font, cp, nm) in enumerate(arms()):
        if cp is None:
            continue
        s, b = spans.get(ai), base.get(font)
        print("%-9s U+%04X %-18s %9s"
              % (font, cp, nm[:18],
                 "MISSING" if s is None or b is None else "%.3f" % ((s - b) / REPEAT)))


def read():
    import re
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.ScreenUpdating = False
    d = app.Documents.Open(docx(), ReadOnly=True)
    ys = {}
    try:
        d.Repaginate()
        for p in d.Paragraphs:
            rng = p.Range
            m = re.match(r"M(\d\d)([SE])", rng.Text)
            if not m:
                continue
            c = d.Range(rng.Start, rng.Start)
            ys[(int(m.group(1)), m.group(2))] = (c.Information(3), round(c.Information(6), 2))
    finally:
        d.Close(False)
        app.Quit()
    spans = {}
    for ai in range(0, len(arms())):
        s, e = ys.get((ai, "S")), ys.get((ai, "E"))
        if s and e and s[0] == e[0]:
            spans[ai] = e[1] - s[1]
    report(spans, "WORD")


def pdf():
    """Which font does Word actually draw each symbol in?

    The line height follows the font Word falls back to when the run's ascii
    font has no glyph, so the fallback has to be identified, not guessed:
    ExportAsFixedFormat, then read the span font name per arm with fitz.
    """
    import fitz
    import win32com.client as w
    out = os.path.join(OUT, os.path.basename(docx()).replace(".docx", ".pdf"))
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx(), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(out, 17)          # wdExportFormatPDF
    finally:
        d.Close(False)
        app.Quit()
    doc = fitz.open(out)
    print("%-9s %-8s %-18s %-26s %s" % ("font", "cp", "name", "Word span font", "size"))
    for ai, (font, cp, nm) in enumerate(arms()):
        if cp is None:
            continue
        page = doc[ai]                          # one arm per page (pageBreakBefore)
        hit = None
        for blk in page.get_text("rawdict")["blocks"]:
            for ln in blk.get("lines", []):
                for sp in ln.get("spans", []):
                    if any(c["c"] == chr(cp) for c in sp.get("chars", [])):
                        hit = (sp["font"], round(sp["size"], 2))
                        break
                if hit:
                    break
            if hit:
                break
        print("%-9s U+%04X %-18s %-26s %s"
              % (font, cp, nm[:18], hit[0] if hit else "(not drawn as text)",
                 hit[1] if hit else ""))


def oxi(envs=""):
    import json
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "symline_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "sym"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    ys = {}
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        for e in pg["elements"]:
            t = (e.get("text") or "").strip()
            if len(t) == 4 and t.startswith("M") and t[3] in "SE" and t[1:3].isdigit():
                ys.setdefault((int(t[1:3]), t[3]), (pg["page"], e["y"]))
    spans = {}
    for ai in range(0, len(arms())):
        s, e = ys.get((ai, "S")), ys.get((ai, "E"))
        if s and e and s[0] == e[0]:
            spans[ai] = e[1] - s[1]
    report(spans, "OXI  " + (envs or "(default)"))


if __name__ == "__main__":
    THEME = "theme" in sys.argv[2:]
    GRID = "grid" in sys.argv[2:]
    CELL = "cell" in sys.argv[2:]
    WITH_CJK = "cjk" in sys.argv[2:]
    if sys.argv[1] == "oxi":
        oxi(next((a for a in sys.argv[2:] if a not in ("theme", "grid", "cell", "cjk")), ""))
    else:
        {"gen": gen, "read": read, "pdf": pdf}[sys.argv[1]]()
