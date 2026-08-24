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
R_TW = list(range(700, 1201, 5))               # 35..60pt in 0.25pt steps
# (the natural line needs r <= 47.2pt, so the flip lives in this window;
#  the first sweep ran 0..20pt and every arm simply held everything.)
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')

# FACE=<font name> puts the arms in that face. The default (unset) inherits the
# base document's ＭＳ 明朝 -- MONOSPACED. S1207 found a proportional Japanese
# face gets no 約物 pool at all; this lets the same arms ask whether the line-end
# HANG dies with it.
FACE = os.environ.get("FACE") or ""
RPR = ('<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
       % (FACE, FACE, FACE)) if FACE else '<w:rFonts w:hint="eastAsia"/>'

ARMS = {
    "NOMARK":   "甲" + "亜" * (NCH - 1),
    "TAILPAIR": "甲" + "亜" * (NCH - 2) + "。",
    "MIDMARK":  "甲" + "亜" * 15 + "、" + "亜" * (NCH - 18) + "。",
}


def build():
    os.makedirs(OUT, exist_ok=True)
    paras, arms = [], []
    for name, txt in ARMS.items():
        assert len(txt) == NCH, (name, len(txt))
        for r in R_TW:
            arms.append((name, r))
            paras.append(
                '<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="both"/>'
                '<w:ind w:left="0" w:right="%d"/>'
                '<w:rPr>%s</w:rPr></w:pPr>'
                '<w:r><w:rPr>%s</w:rPr>'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                % (r, RPR, RPR, txt))
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    body = "<w:body>" + "".join(paras) + sect + "</w:body>"
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s>%s</w:document>' % (W_NS, body))
    dst = os.path.join(OUT, "poolorder%s.docx" % ("_" + FACE if FACE else ""))
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    with open(os.path.join(OUT, "arms%s.txt" % ("_" + FACE if FACE else "")),
              "w", encoding="utf-8") as fh:
        for a in arms:
            fh.write("%s\t%d\n" % a)
    print("built %s (%d paragraphs)" % (dst, len(paras)))


def to_pdf():
    import win32com.client as wc
    tag = "_" + FACE if FACE else ""
    docx = os.path.join(OUT, "poolorder%s.docx" % tag)
    pdf = os.path.join(OUT, "poolorder%s.pdf" % tag)
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
    arms = [l.split("\t") for l in open(os.path.join(OUT, "arms.txt"),
                                        encoding="utf-8").read().splitlines()]
    doc = fitz.open(os.path.join(OUT, "poolorder.pdf"))
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
        for _, ch in sorted(rows):
            if ch[0]["c"] == "甲":
                heads.append(len(ch))
    if len(heads) != len(arms):
        print("%d paragraph heads for %d arms -- aborting" % (len(heads), len(arms)))
        return
    by = {}
    for (name, r), n in zip(arms, heads):
        by.setdefault(name, []).append((int(r) / 20.0, n))
    print("NCH=%d  em=%.2f  content=%.1fpt  face=%s"
          % (NCH, EM, CONTENT_PT, FACE or "(inherited MS Mincho)"))
    print("   arm        last r holding all %d   natural needs r <=   credit (em)" % NCH)
    for name in ("NOMARK", "TAILPAIR", "MIDMARK"):
        rows = by[name]
        top = max(n for _, n in rows)      # Word's PDF adds a trailing space glyph
        full = [r for r, n in rows if n >= top]
        if not full:
            print("   %-9s never fits" % name)
            continue
        r_flip = max(full)
        natural_r = CONTENT_PT - NCH * EM
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
