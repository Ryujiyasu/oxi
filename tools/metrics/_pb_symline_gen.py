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

  python _pb_symline_gen.py gen
  python _pb_symline_gen.py read              # Word COM truth
  python _pb_symline_gen.py oxi               # Oxi, same arms
"""
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_symline")
DOCX = os.path.join(OUT, "symline.docx")
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
    (0x3007, "IDEOGRAPHIC ZERO"),
    (0x4E00, "CJK control"),
]
FONTS = ["Arial", "Calibri"]


def arms():
    return [(f, cp, nm) for f in FONTS for cp, nm in CHARS]


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
    body = [marker("M00S", FONTS[0], True), marker("M00E", FONTS[0])]
    for ai, (font, cp, _nm) in enumerate(arms(), start=1):
        body.append(marker("M%02dS" % ai, font, True))
        for _ in range(REPEAT):
            body.append(subject(font, cp))
        body.append(marker("M%02dE" % ai, font))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>')
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(arms()), "arms x", REPEAT)


def report(spans, who):
    base = spans.get(0)
    if base is None:
        raise SystemExit("control arm missing")
    print("%s   marker-only span = %.2f" % (who, base))
    print("%-9s %-8s %-18s %9s" % ("font", "cp", "name", "per copy"))
    for ai, (font, cp, nm) in enumerate(arms(), start=1):
        s = spans.get(ai)
        print("%-9s U+%04X %-18s %9s"
              % (font, cp, nm[:18], "MISSING" if s is None else "%.3f" % ((s - base) / REPEAT)))


def read():
    import re
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.ScreenUpdating = False
    d = app.Documents.Open(DOCX, ReadOnly=True)
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
    for ai in range(0, len(arms()) + 1):
        s, e = ys.get((ai, "S")), ys.get((ai, "E"))
        if s and e and s[0] == e[0]:
            spans[ai] = e[1] - s[1]
    report(spans, "WORD")


def oxi(envs=""):
    import json
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "symline_oxi.json")
    subprocess.run([GDI, DOCX, os.path.join(tempfile.gettempdir(), "sym"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    ys = {}
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        for e in pg["elements"]:
            t = (e.get("text") or "").strip()
            if len(t) == 4 and t.startswith("M") and t[3] in "SE" and t[1:3].isdigit():
                ys.setdefault((int(t[1:3]), t[3]), (pg["page"], e["y"]))
    spans = {}
    for ai in range(0, len(arms()) + 1):
        s, e = ys.get((ai, "S")), ys.get((ai, "E"))
        if s and e and s[0] == e[0]:
            spans[ai] = e[1] - s[1]
    report(spans, "OXI  " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "read": read}[sys.argv[1]]()
