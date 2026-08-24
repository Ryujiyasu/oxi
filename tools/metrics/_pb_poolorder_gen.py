# -*- coding: utf-8 -*-
"""Can a line borrow aki from a mark that is not on it YET?

S1209 fixed the SIZE of the 約物 pool (half an em, released to 1.5 by a fullwidth
space or an opening bracket before the squeeze). Its arms always placed the marks
BEFORE the character being squeezed in, so one question never came up: the
implementation sums `current_line_chars.chain(buf_chars)`, i.e. it also counts a
mark that arrives WITH the character it is paying for.

d77a58485f16 p9 is that case. Word breaks

    カ 本利用ルールは…適用されることもありま | す。

but Oxi with OXI_YAKUCOMP keeps す。 on line 1 by compressing 3.9pt -- and the
ONLY mark on that line is the 。 at the very end, the character being squeezed.

So: sweep the right indent through the region where the SECOND-TO-LAST character
overflows, with the mark sitting after it, and see whether Word admits the pair.

    TAILPAIR ... 亜 亜 。   the overflowing char is 亜, a 。 follows it
    MIDMARK  ... 、 亜 亜 。 the same, but a 、 sits earlier on the line
    NOMARK   ... 亜 亜 亜   control, no mark anywhere

Prediction stated before measuring: TAILPAIR behaves like NOMARK (no credit --
a mark cannot lend what it has not yet been placed to lend) and MIDMARK gets the
half em.

    python _pb_poolorder_gen.py gen
    python _pb_poolorder_gen.py pdf
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_poolorder")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
CONTENT_PT = (11906 - 1701 - 1701) / 20.0      # 425.2
EM = 10.5
NCH = 36
# The flip lives at r = measure - NCH*em; centre the window on it (the first
# sweep ran a fixed 0..20pt and every arm simply held everything, and a fixed
# 35..60pt window then missed the CELL flip entirely because the cell's measure
# is 25pt narrower).
def _window():
    measure = (CELL_TW / 20.0) if CELL else CONTENT_PT
    mid = measure - NCH * EM - (21.2 if HANG else 0.0)
    return list(range(max(0, int((mid - 12) * 20)), int((mid + 13) * 20), 5))
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')

# FACE=<font name> puts the arms in that face. The default (unset) inherits the
# base document's ＭＳ 明朝 -- MONOSPACED. S1207 found a proportional Japanese
# face gets no 約物 pool at all; this lets the same arms ask whether the line-end
# HANG dies with it.
FACE = os.environ.get("FACE") or ""
# COMPAT15=1 rewrites the base document's settings.xml to the modern engine:
# compatibilityMode 15 and NO <w:useAltKinsokuLineBreakRules/>. The base
# (tokyoshugyo) is compat 11 WITH the alternate rules -- and it is the ONLY
# compat-11 document in the corpus; every other one, d77a58 included, is 15.
# Every 約物 rule so far was derived on the compat-11 engine.
COMPAT15 = os.environ.get("COMPAT15") == "1"
# CELL=1 puts every arm in a fixed-width one-cell table (margins zeroed) instead
# of the body. d77a58's line is a CELL line in a proportional face, and Word does
# not hang its line-final 。 there while the body arms hang theirs whole.
CELL = os.environ.get("CELL") == "1"
CELL_TW = 8000                                  # 400pt, wide enough for 36 chars
# HANG=1 reproduces d77a58's paragraph shape: a hanging indent with a marker and
# a tab in front of the text (left 564tw, hanging 140tw). Everything else about
# that line has now been matched -- cell, proportional face, jc=both, compat --
# and Word still refuses the 0.78pt its 。 needs, so this is the last difference.
# HANG=1 indent + marker + tab (d77a58's shape); HANG=2 the indent alone, to
# tell which of the two kills the hang.
HANG = os.environ.get("HANG") in ("1", "2")
HANG_TAB = os.environ.get("HANG") == "1"
IND = '<w:ind w:left="564" w:hanging="140"/>' if HANG else '<w:ind w:left="0" w:right="%d"/>'
PREFIX = "カ	" if HANG_TAB else ""
RPR = ('<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
       % (FACE, FACE, FACE)) if FACE else '<w:rFonts w:hint="eastAsia"/>'

ARMS = {
    "NOMARK":   "甲" + "亜" * (NCH - 1),
    "TAILPAIR": "甲" + "亜" * (NCH - 2) + "。",
    "MIDMARK":  "甲" + "亜" * 15 + "、" + "亜" * (NCH - 18) + "。",
}


def tag():
    return (("_" + FACE if FACE else "") + ("_c15" if COMPAT15 else "")
            + ("_cell" if CELL else "")
            + (("_hang" if HANG_TAB else "_ind") if HANG else ""))


def wrap_cell(para):
    """One fixed-width cell with zero margins around the paragraph."""
    return ('<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/>'
            '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
            '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
            '</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc>'
            '</w:tr></w:tbl>'
            '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>'
            % (CELL_TW, CELL_TW, CELL_TW, para))


def modern(settings):
    """The compat-11 base turned into the engine the corpus actually uses."""
    settings = settings.replace("<w:useAltKinsokuLineBreakRules/>", "")
    return re.sub(r'(w:name="compatibilityMode"[^>]*w:val=")\d+',
                  r"\g<1>15", settings)


def build():
    os.makedirs(OUT, exist_ok=True)
    paras, arms = [], []
    for name, txt in ARMS.items():
        assert len(txt) == NCH, (name, len(txt))
        for r in _window():
            arms.append((name, r))
            ind = ('<w:ind w:left="564" w:hanging="140" w:right="%d"/>' % r
                   if HANG else '<w:ind w:left="0" w:right="%d"/>' % r)
            para = ('<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="both"/>'
                    + ind
                    + ('<w:rPr>%s</w:rPr></w:pPr>'
                       '<w:r><w:rPr>%s</w:rPr>'
                       '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                       % (RPR, RPR, PREFIX + txt)))
            paras.append(wrap_cell(para) if CELL else para)
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    body = "<w:body>" + "".join(paras) + sect + "</w:body>"
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s>%s</w:document>' % (W_NS, body))
    dst = os.path.join(OUT, "poolorder%s.docx" % tag())
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        elif COMPAT15 and item.filename == "word/settings.xml":
            data = modern(data.decode("utf-8")).encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    with open(os.path.join(OUT, "arms%s.txt" % tag()), "w", encoding="utf-8") as fh:
        for a in arms:
            fh.write("%s\t%d\n" % a)
    print("built %s (%d paragraphs)" % (dst, len(paras)))


def to_pdf():
    import win32com.client as wc
    docx = os.path.join(OUT, "poolorder%s.docx" % tag())
    pdf = os.path.join(OUT, "poolorder%s.pdf" % tag())
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf),
                              ExportFormat=17, OpenAfterExport=False)
        d.Close(False)
    finally:
        app.Quit()
    return pdf


def measure():
    import fitz
    arms = [l.split("\t") for l in open(os.path.join(OUT, "arms%s.txt" % tag()),
                                        encoding="utf-8").read().splitlines()]
    doc = fitz.open(os.path.join(OUT, "poolorder%s.pdf" % tag()))
    heads = []
    for page in doc:
        rows = []
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                ch = sorted([c for s in l["spans"] for c in s["chars"]],
                            key=lambda c: c["origin"][0])
                if ch:
                    rows.append((round(l["bbox"][1], 1), ch))
        for _, ch in sorted(rows, key=lambda t: t[0]):
            if ch[0]["c"] == "甲":
                heads.append(len(ch))
    if len(heads) != len(arms):
        print("%d paragraph heads for %d arms -- aborting" % (len(heads), len(arms)))
        return
    by = {}
    for (name, r), n in zip(arms, heads):
        by.setdefault(name, []).append((int(r) / 20.0, n))
    print("NCH=%d  em=%.2f  measure=%.1fpt  face=%s  engine=%s  where=%s"
          % (NCH, EM, CELL_TW / 20.0 if CELL else CONTENT_PT,
             FACE or "(inherited MS Mincho)",
             "compat 15, no altKinsoku" if COMPAT15 else "the base's compat 11 + altKinsoku",
             ("CELL" if CELL else "body") + ((" + hanging indent and a tab" if HANG_TAB else " + hanging indent")
              if HANG else "")))
    print("   arm        last r holding all %d   natural needs r <=   credit (em)" % NCH)
    for name in ("NOMARK", "TAILPAIR", "MIDMARK"):
        rows = by[name]
        top = max(n for _, n in rows)      # Word's PDF adds a trailing space glyph
        full = [r for r, n in rows if n >= top]
        if not full:
            print("   %-9s never fits" % name)
            continue
        r_flip = max(full)
        natural_r = (CELL_TW / 20.0 if CELL else CONTENT_PT) - NCH * EM
        print("   %-9s %8.2f            %8.2f          %+.3f   (top n=%d)"
              % (name, r_flip, natural_r, (r_flip - natural_r) / EM, top))
    for name in ("NOMARK", "TAILPAIR", "MIDMARK"):
        rows = sorted(by[name])
        counts = {}
        for r, n in rows:
            counts.setdefault(n, []).append(r)
        print("   %-9s counts: %s" % (name, {k: (min(v), max(v)) for k, v in
                                             sorted(counts.items(), reverse=True)}))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "gen":
        build()
    elif cmd == "pdf":
        to_pdf()
        measure()
    else:
        measure()
