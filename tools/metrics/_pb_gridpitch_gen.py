# -*- coding: utf-8 -*-
"""What advance does a docGrid give a fullwidth character, at every font size?

Declared BEFORE measuring:

    additive      adv = fs + charSpace/4096          <- the claim
    proportional  adv = fs * (10.5 + charSpace/4096) / 10.5   <- what Oxi used
    natural       adv = fs                           <- what Oxi uses in CELLS,
                                                        and what S141 claims for
                                                        fs < the grid default

An arm is one paragraph of NCH identical fullwidth characters (no 約物, so no
compression credit can pay for an overflow) at font size `fs`, left aligned so
nothing is stretched, with the paragraph's RIGHT INDENT swept in 0.25pt steps.
The largest indent that still leaves all NCH characters on line 1 brackets the
advance to 0.25/NCH pt -- fine enough to separate the three predictions, whose
spread at fs=9 is 0.05pt/char.

MEASURED 2026-08-24 (charSpace 1966 / 1453 / 532 / -2714 / no char grid, sizes
9 / 10 / 10.5 / 12, right indent 0..30pt in 0.25pt steps = 2420 arms):

    advance   = fs + charSpace/4096          additive, size-independent, both signs
    line width= floor(content/pitch) * pitch  pitch = default_fs + charSpace/4096

18 of the 20 (charSpace, size) pairs land in the predicted 0.25pt bracket; the
two that miss are both fs=9 and both by less than 0.03pt over 44-49 characters.
The PROPORTIONAL form (fs * pitch / default_fs) is out by 0.05pt/char at fs 9
and fs 12 -- 2pt over a line -- and the flooring is what made the corpus look
like it wanted a "stretched" pitch: the pitch is raw, the WIDTH is truncated.

    python _pb_gridpitch_gen.py gen
    python _pb_gridpitch_gen.py pdf
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_gridpitch")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
CONTENT_PT = (11906 - 1701 - 1701) / 20.0        # 425.2pt
DEFAULT_FS = 10.5                                 # Normal sz=21, no rPrDefault sz
CHAR_SPACES = [1966, -2714, 1453, 532, 0]
SIZES = [9.0, 10.0, 10.5, 12.0]
R_TW = list(range(0, 601, 5))                     # 0..30pt in 0.25pt steps
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')


def nch_for(fs, cs=0):
    """Enough characters that the flip lands inside the swept indent range.

    The count must follow the ADVANCE, not the font size: with charSpace -2714
    a 9pt character advances 8.34pt, and a count picked off 9.00 puts the flip
    at r = 48pt, past the end of the sweep (those arms read "never fits").
    """
    return int((CONTENT_PT - 12.0) / (fs + cs / 4096.0))


def build_one(cs):
    paras, arms = [], []
    for fs in SIZES:
        n = nch_for(fs, cs)
        txt = "甲" + "亜" * (n - 1)
        for r in R_TW:
            arms.append((cs, fs, n, r))
            paras.append(
                '<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="left"/>'
                '<w:ind w:left="0" w:right="%d"/>'
                '<w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr>'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                % (r, int(fs * 2), int(fs * 2), txt))
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    grid = ('<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="%d"/>' % cs
            if cs else '<w:docGrid w:type="lines" w:linePitch="360"/>')
    sect = re.sub(r"<w:docGrid[^>]*/>", grid, sect)
    body = "<w:body>" + "".join(paras) + sect + "</w:body>"
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s>%s</w:document>' % (W_NS, body))
    dst = os.path.join(OUT, "cs%d.docx" % cs)
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    with open(os.path.join(OUT, "cs%d.arms" % cs), "w", encoding="utf-8") as fh:
        for a in arms:
            fh.write("%d\t%.1f\t%d\t%d\n" % a)
    print("built %s (%d paragraphs)" % (dst, len(paras)))


def build():
    os.makedirs(OUT, exist_ok=True)
    for cs in CHAR_SPACES:
        build_one(cs)


def to_pdf():
    import win32com.client as wc
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    try:
        for cs in CHAR_SPACES:
            docx = os.path.join(OUT, "cs%d.docx" % cs)
            pdf = os.path.join(OUT, "cs%d.pdf" % cs)
            d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf),
                                  ExportFormat=17, OpenAfterExport=False)
            d.Close(False)
            print("exported", pdf)
    finally:
        app.Quit()


def first_line_counts(pdf):
    """Per paragraph (a line starting with 甲), how many glyphs are on line 1."""
    import fitz
    doc = fitz.open(pdf)
    out = []
    for page in doc:
        rows = []
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                ch = [c for s in l["spans"] for c in s["chars"]]
                if ch:
                    rows.append((round(l["bbox"][1], 1), ch))
        for _, ch in sorted(rows):
            first = ch[0]["c"]
            if first == "甲" or first == "�":   # paragraph head
                out.append(len(ch))
            elif out:
                pass                                  # continuation line
    return out


def measure():
    for cs in CHAR_SPACES:
        arms = [l.split("\t") for l in open(os.path.join(OUT, "cs%d.arms" % cs),
                                            encoding="utf-8").read().splitlines()]
        counts = first_line_counts(os.path.join(OUT, "cs%d.pdf" % cs))
        if len(counts) != len(arms):
            print("cs=%d: %d paragraph heads for %d arms -- skipped"
                  % (cs, len(counts), len(arms)))
            continue
        cs_pt = cs / 4096.0
        print("\ncharSpace=%d (%+.4f pt/char)   content=%.1fpt" % (cs, cs_pt, CONTENT_PT))
        print("   fs    NCH  r_flip   measured adv      additive  proportional   natural")
        for fs in SIZES:
            n = nch_for(fs, cs)
            rows = [(int(a[3]), c) for a, c in zip(arms, counts)
                    if abs(float(a[1]) - fs) < 1e-6]
            full = [r for r, c in rows if c >= n]
            if not full:
                print("   %5.1f  %3d   never fits" % (fs, n))
                continue
            r_flip = max(full) / 20.0
            hi = (CONTENT_PT - r_flip) / n
            lo = (CONTENT_PT - r_flip - 0.25) / n
            add = fs + cs_pt
            prop = fs * (DEFAULT_FS + cs_pt) / DEFAULT_FS
            # The line's usable width, if Word truncates it to a whole number
            # of grid cells at the DEFAULT size's pitch.
            pitch = DEFAULT_FS + cs_pt
            usable = (int(CONTENT_PT / pitch) * pitch) if cs else CONTENT_PT
            need = n * add
            fits = need <= usable - r_flip
            breaks = need > usable - r_flip - 0.25
            print("   %5.1f  %3d  %6.2f   %7.4f..%7.4f  %8.4f  %12.4f  %8.4f   "
                  "floor: need %.2f in %.2f -> %s"
                  % (fs, n, r_flip, lo, hi, add, prop, fs, need, usable - r_flip,
                     "OK" if (fits and breaks) else "MISS"))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "gen":
        build()
    elif cmd == "pdf":
        to_pdf()
        measure()
    else:
        measure()
