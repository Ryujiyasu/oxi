# -*- coding: utf-8 -*-
"""Does w:textAlignment change what a Symbol bullet adds to a line's height?

educational__005c3a80's "Body copy" list (Arial 12 text, Symbol 12 bullet,
single spacing, compat 15, style textAlignment=center) advances 14.04pt per
item in Word's PDF, while the S820b model (Symbol asc + Arial desc = the
_pb_bullet_gen overflow model) gives 14.61 and a plain Arial 12 line is 13.80.
Oxi follows S820b, so the list drifts +0.46pt per bullet (census 2.77 -> 6.9).

Each arm is NRUN single-line bulleted paragraphs in its own section; the span
between the first and last on the page pins the per-item height (Info6 is the
96dpi pixel, so +-0.75/(n-1)).

  python _pb_txalign_bullet_gen.py gen
  python _pb_txalign_bullet_gen.py read
"""
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_bullet_gen as B  # noqa: E402

OUT = os.path.join(B.REPO, "pipeline_data", "_pb_txalign_bullet")
DOCX = os.path.join(OUT, "txalign_bullet.docx")
NRUN = 20

SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + B.NS + '>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
CT = B.CT.replace('</Types>',
                  '<Override PartName="/word/settings.xml" ContentType="application/'
                  'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
DRELS = B.DRELS.replace('</Relationships>',
                        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/'
                        'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                        '</Relationships>')
# "Body copy" as in the source document (keepNext, autoSpace off, 10pt bold
# style run properties that the paragraph's runs override to Arial 12 regular).
STYLES = B.STYLES.replace(
    '</w:styles>',
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Bodycopy"><w:name w:val="Body copy"/>'
    '<w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:autoSpaceDE w:val="0"/>'
    '<w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/><w:jc w:val="left"/>'
    '<w:textAlignment w:val="center"/></w:pPr><w:rPr><w:rFonts w:ascii="Helvetica Now Text Light" '
    'w:hAnsi="Helvetica Now Text Light" w:cs="Minion Pro"/><w:b/><w:bCs/><w:color w:val="000000"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style></w:styles>')

# (label, textAlignment or None, bullet?, pStyle or None[, text font, text hsz, marker hsz])
ARMS = [
    ("none",     None,       True,  None),
    ("auto",     "auto",     True,  None),
    ("baseline", "baseline", True,  None),
    ("center",   "center",   True,  None),
    ("top",      "top",      True,  None),
    ("bottom",   "bottom",   True,  None),
    ("bodycopy", None,       True,  "Bodycopy"),   # faithful slice of the doc
    ("plainctr", "center",   False, None),         # no bullet, centre
    ("plain",    None,       False, None),
    # centre law across faces / sizes / a marker larger or smaller than the text
    ("cAr10",    "center",   True,  None, "Arial", 20, None),
    ("cAr12m16", "center",   True,  None, "Arial", 24, 32),
    ("cAr12m9",  "center",   True,  None, "Arial", 24, 18),
    ("cCal11",   "center",   True,  None, "Calibri", 22, None),
    ("cCal14",   "center",   True,  None, "Calibri", 28, None),
    ("cTnr12",   "center",   True,  None, "Times New Roman", 24, None),
    ("bTnr12",   "baseline", True,  None, "Times New Roman", 24, None),
    ("pTnr12",   None,       False, None, "Times New Roman", 24, None),
    ("pCal11",   None,       False, None, "Calibri", 22, None),
    ("pCal14",   None,       False, None, "Calibri", 28, None),
    ("pAr10",    None,       False, None, "Arial", 20, None),
]


def numbering():
    lv = lambda i, sz: (
        '<w:abstractNum w:abstractNumId="%d"><w:multiLevelType w:val="hybridMultilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val=""/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>%s</w:rPr></w:lvl>'
        '</w:abstractNum>' % (i, ('<w:sz w:val="%d"/>' % sz) if sz else ""))
    sizes = [None] + sorted({a[6] for a in ARMS if len(a) > 6 and a[6]})
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + B.NS + '>'
            + "".join(lv(i, sz) for i, sz in enumerate(sizes))
            + "".join('<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>' % (i + 1, i)
                      for i in range(len(sizes))) + '</w:numbering>')


def num_id(msz):
    sizes = [None] + sorted({a[6] for a in ARMS if len(a) > 6 and a[6]})
    return sizes.index(msz) + 1


def _numbering_single():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + B.NS + '>'
            '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
            '<w:lvlText w:val=""/><w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
            '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>'
            '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')


def para(tag, ta, bullet, pstyle, font="Arial", hsz=24, msz=None):
    rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/><w:b w:val="0"/>'
           '<w:bCs w:val="0"/><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (font, font, font, hsz, hsz))
    ppr = ""
    if pstyle:
        ppr += '<w:pStyle w:val="%s"/>' % pstyle
    if bullet:
        ppr += '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>' % num_id(msz)
    ppr += '<w:widowControl w:val="0"/>'
    if ta:
        ppr += '<w:textAlignment w:val="%s"/>' % ta
    return '<w:p><w:pPr>%s%s</w:pPr><w:r>%s<w:t>%s</w:t></w:r></w:p>' % (ppr, rpr, rpr, tag)


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, arm in enumerate(ARMS):
        for k in range(NRUN):
            body.append(para("T%02dK%02d" % (ai, k), *arm[1:]))
        body.append(B.sect(ai == len(ARMS) - 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + B.NS +
           '><w:body>' + "".join(body) + "</w:body></w:document>")
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", B.RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/numbering.xml", numbering())
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(ARMS) * NRUN, "paragraphs")


def read():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(DOCX, ReadOnly=True)
    rows = {}
    try:
        d.Repaginate()
        for i in range(1, d.Paragraphs.Count + 1):
            rng = d.Paragraphs(i).Range
            m = re.match(r"(T\d+K\d+)", rng.Text)
            if not m:
                continue
            c = d.Range(rng.Start, rng.Start)
            rows[m.group(1)] = (c.Information(3), round(c.Information(6), 2))
    finally:
        d.Close(False)
        app.Quit()
    print("%-10s %-9s %-6s %9s %7s" % ("arm", "txAlign", "bullet", "per-item", "+-"))
    for ai, (label, ta, bullet, _ps, *_rest) in enumerate(ARMS):
        pts = [rows.get("T%02dK%02d" % (ai, k)) for k in range(NRUN)]
        if not pts[0]:
            continue
        pg0 = pts[0][0]
        last = max(k for k in range(NRUN) if pts[k] and pts[k][0] == pg0)
        print("%-10s %-9s %-6s %9.4f %7.3f"
              % (label, ta or "-", bullet, (pts[last][1] - pts[0][1]) / last, 0.75 / last))


def pdf():
    """Exact per-item pitch from Word's PDF baselines (no 96dpi quantum)."""
    import fitz
    import win32com.client as w
    out = os.path.join(OUT, "txalign_bullet.pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(DOCX, ReadOnly=True)
        d.SaveAs2(out, 17)
        d.Close(False)
    finally:
        app.Quit()
    base = {}
    for pno, page in enumerate(fitz.open(out)):
        for b in page.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                for s in ln["spans"]:
                    t = "".join(c["c"] for c in s["chars"])
                    m = re.search(r"T(\d+)K(\d+)", t)
                    if m:
                        base[(int(m.group(1)), int(m.group(2)))] = (pno, s["chars"][0]["origin"][1])
    print("%-10s %-9s %-6s %9s %9s %7s" % ("arm", "txAlign", "bullet", "per-item", "model", "err"))
    for ai, (label, ta, bullet, _ps, *rest) in enumerate(ARMS):
        pts = [base.get((ai, k)) for k in range(NRUN)]
        if not pts[0]:
            continue
        last = max(k for k in range(NRUN) if pts[k] and pts[k][0] == pts[0][0])
        h = (pts[last][1] - pts[0][1]) / last
        font, hsz, msz = (rest + [None] * 3)[:3] if rest else ("Arial", 24, None)
        mdl = model(ta, bullet, font or "Arial", (hsz or 24) / 2.0, (msz or hsz or 24) / 2.0)
        print("%-10s %-9s %-6s %9.4f %9.4f %+7.3f" % (label, ta or "-", bullet, h, mdl, h - mdl))


FILES = {"Arial": "arial.ttf", "Calibri": "calibri.ttf", "Times New Roman": "times.ttf",
         "Symbol": "symbol.ttf"}


def metrics(face):
    """(win asc, win desc, ext) in em; ext = hhea sum beyond the win box."""
    from fontTools.ttLib import TTFont
    f = TTFont(os.path.join(os.environ["WINDIR"], "Fonts", FILES[face]), lazy=True)
    u = f["head"].unitsPerEm; o = f["OS/2"]; h = f["hhea"]
    wa, wd = o.usWinAscent / u, o.usWinDescent / u
    return wa, wd, max(0.0, (h.ascender - h.descender + h.lineGap) / u - wa - wd)


def model(ta, bullet, face, ts, ms):
    """Line height of a single-spaced line of `face` at ts with a Symbol bullet at ms.
    The marker adds its ASCENT only (S820b); where it sits against the text:
      baseline/auto/none: same baseline       top: no growth
      center: its (asc-desc)/2 midpoint on the text's -> baseline shifted down by
              ((A-D) - (a-d)) / 2"""
    a, d, g = (x * ts for x in metrics(face))
    plain = a + d + g
    if not bullet or ta == "top":
        return plain
    A, D = (x * ms for x in metrics("Symbol")[:2])
    off = ((A - D) - (a - d)) / 2.0 if ta == "center" else 0.0
    return plain + max(0.0, A - off - (a + g))


if __name__ == "__main__":
    {"gen": gen, "read": read, "pdf": pdf}[sys.argv[1]]()
