# -*- coding: utf-8 -*-
"""The same grid-pitch question, asked inside a TABLE CELL.

`_pb_gridpitch_gen.py` settled the body: a fullwidth character advances by
fs + charSpace/4096 whatever its size, and the BODY line's usable width is
truncated to a whole number of grid cells. Oxi's cell breaker is a separate
greedy wrapper, and a1d6e4ef's note column says the cell behaves differently,
so ask the cell the same two questions directly.

Each arm is one paragraph of NCH fullwidth characters inside a fixed-width
one-cell table (cell margins zeroed, jc=left so nothing is stretched), with the
paragraph's RIGHT INDENT swept in 0.25pt steps.

MEASURED 2026-08-24 (3 charSpace x 3 sizes x 81 indent steps = 729 arms): the
cell takes the SAME additive advance (fs + charSpace/4096, fs=9 included) and
does NOT truncate its width to whole grid cells -- all 9 break widths sit within
0.25pt of n * (fs + charSpace/4096) against the cell's own inner width. That
asymmetry with the body is the finding: pitch is shared, flooring is not.

    python _pb_cellpitch_gen.py gen
    python _pb_cellpitch_gen.py pdf
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cellpitch")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
CELL_TW = 4000                      # 200.0pt, margins zeroed
CELL_PT = CELL_TW / 20.0
DEFAULT_FS = 10.5
CHAR_SPACES = [1453, 532, 0]
SIZES = [9.0, 10.5, 12.0]
R_TW = list(range(0, 401, 5))       # 0..20pt in 0.25pt steps
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')


def nch_for(fs):
    return int((CELL_PT - 10.0) / fs)


def table_for(fs, r, txt):
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
        '</w:tblCellMar></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="left"/>'
        '<w:ind w:left="0" w:right="%d"/>'
        '<w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr>'
        '<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
        '<w:p><w:pPr><w:pStyle w:val="a"/><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>'
        % (CELL_TW, CELL_TW, CELL_TW, r, int(fs * 2), int(fs * 2), txt))


def build_one(cs):
    blocks, arms = [], []
    for fs in SIZES:
        n = nch_for(fs)
        txt = "甲" + "亜" * (n - 1)
        for r in R_TW:
            arms.append((cs, fs, n, r))
            blocks.append(table_for(fs, r, txt))
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    grid = ('<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="%d"/>' % cs
            if cs else '<w:docGrid w:type="lines" w:linePitch="360"/>')
    sect = re.sub(r"<w:docGrid[^>]*/>", grid, sect)
    body = "<w:body>" + "".join(blocks) + sect + "</w:body>"
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s>%s</w:document>' % (W_NS, body))
    dst = os.path.join(OUT, "cell%d.docx" % cs)
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    with open(os.path.join(OUT, "cell%d.arms" % cs), "w", encoding="utf-8") as fh:
        for a in arms:
            fh.write("%d\t%.1f\t%d\t%d\n" % a)
    print("built %s (%d tables)" % (dst, len(blocks)))


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
            docx = os.path.join(OUT, "cell%d.docx" % cs)
            pdf = os.path.join(OUT, "cell%d.pdf" % cs)
            d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf),
                                  ExportFormat=17, OpenAfterExport=False)
            d.Close(False)
            print("exported", pdf)
    finally:
        app.Quit()


def heads(pdf):
    """(glyphs on the paragraph's first line, advance on it) per arm."""
    import fitz
    doc = fitz.open(pdf)
    out = []
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
            # ONE row per arm: a paragraph head starts with 甲, continuation
            # lines start with 亜. Indexing every line instead shifts the whole
            # mapping as soon as one arm wraps (it did, and the fs=10.5 rows
            # then read as if a 188pt line held 18 characters).
            if len(ch) < 3 or ch[0]["c"] != "甲":
                continue
            adv = (ch[-1]["origin"][0] - ch[0]["origin"][0]) / (len(ch) - 1)
            out.append((len(ch), adv, ch[0]["origin"][0]))
    return out


def measure():
    for cs in CHAR_SPACES:
        arms = [l.split("\t") for l in open(os.path.join(OUT, "cell%d.arms" % cs),
                                            encoding="utf-8").read().splitlines()]
        rows = heads(os.path.join(OUT, "cell%d.pdf" % cs))
        if len(rows) < len(arms):
            print("cell cs=%d: %d lines for %d arms -- continuation lines merged?"
                  % (cs, len(rows), len(arms)))
        cs_pt = cs / 4096.0
        pitch = DEFAULT_FS + cs_pt
        cells = int(CELL_PT / pitch)
        print(chr(10) + "cell charSpace=%d (%+.4f)  cell=%.1fpt  whole cells=%d x %.4f = %.2f"
              % (cs, cs_pt, CELL_PT, cells, pitch, cells * pitch))
        print("   fs   NCH  r_flip   advance   break width   no-floor pred   floor pred")
        k = 0
        for fs in SIZES:
            n = nch_for(fs)
            got = []
            for a in arms:
                if abs(float(a[1]) - fs) > 1e-6:
                    continue
                if k < len(rows):
                    got.append((int(a[3]), rows[k][0], rows[k][1]))
                k += 1
            full = [(r, adv) for r, cnt, adv in got if cnt >= n]
            if not full:
                print("   %5.1f %3d   never fits" % (fs, n))
                continue
            r_flip = max(r for r, _ in full) / 20.0
            adv = [a for r, a in full if r / 20.0 == r_flip][0]
            width = CELL_PT - r_flip
            print("   %5.1f %3d  %6.2f  %8.4f   %8.2f      %8.2f       %8.2f"
                  % (fs, n, r_flip, adv, width, n * (fs + cs_pt),
                     cells * pitch - r_flip))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "gen":
        build()
    elif cmd == "pdf":
        to_pdf()
        measure()
    else:
        measure()
