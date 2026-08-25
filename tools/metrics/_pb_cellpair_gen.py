# -*- coding: utf-8 -*-
"""With the pair compression on, does a CELL break where Word breaks?

S1218 gives the cell breaker the adjacent-mark rule the body has had since S532,
and the widths then match Word on all ten `_pb_yakuwidth_gen.py` arms -- but
tokyoshugyo loses nine paragraphs. Either the break-time capacity already paid
for those pairs, or the regression is elsewhere. So compare the BREAK directly:
one text, in a cell, right indent swept in 0.25pt steps, read from Word's PDF and
from Oxi's own dump.

    python _pb_cellpair_gen.py
"""
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cellpair")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
CELL_TW = 8000
EM = 10.5
TEXTS = {
    "no_pair": "甲" + "亜" * 30 + "、亜。亜亜",
    "one_pair": "甲" + "亜" * 30 + "。）亜亜亜",
    "two_pairs": "甲" + "亜" * 27 + "。）亜。）亜亜",
    # tokyoshugyo's p20 line ends 「…手待ち時間」）」 -- the pair IS the line end.
    # A line-final mark also HANGS (S1205: half an em in a cell), so read these
    # three together: `pair_end` carries hang + (pair?), `one_mark_end` carries
    # the hang alone, and their difference is the pair's contribution.
    "pair_end": "甲" + "亜" * 33 + "。）",
    "one_mark_end": "甲" + "亜" * 34 + "。",
    "pair_then_one": "甲" + "亜" * 32 + "。）亜",
    # tokyoshugyo's pair is TWO CLOSING BRACKETS (「…時間」）」), not period+paren.
    # S1199 found a run of closers behaves differently for the HANG (one closer
    # hangs 441/441, two or more 1/154), so ask the same of the compression.
    "closers_end": "甲" + "亜" * 33 + "」）",
    "closers_mid": "甲" + "亜" * 30 + "」）亜亜亜",
    "quote_pair_end": "甲" + "亜" * 32 + "「亜」）",
}
# the flip sits near 400pt - 36 chars x 10.5 = 22pt, plus whatever the pairs
# save; the first window (0..20pt) ended just short of it and every arm read
# 'still holds everything'.
R_TW = list(range(300, 801, 5))


def build():
    os.makedirs(OUT, exist_ok=True)
    blocks, index = [], []
    for name, txt in TEXTS.items():
        for r in R_TW:
            index.append((name, r, len(txt)))
            blocks.append(
                '<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/>'
                '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
                '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
                '</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
                '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>'
                '<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="both"/>'
                '<w:ind w:left="0" w:right="%d"/>'
                '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
                '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>'
                % (CELL_TW, CELL_TW, CELL_TW, r, txt))
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
           '<w:body>%s%s</w:body></w:document>' % ("".join(blocks), sect))
    dst = os.path.join(OUT, "cellpair.docx")
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    return dst, index


def word_counts(docx):
    import win32com.client as wc
    import fitz
    pdf = os.path.splitext(docx)[0] + ".pdf"
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf),
                              ExportFormat=17, OpenAfterExport=False)
        d.Close(False)
    finally:
        app.Quit()
    out = []
    for page in fitz.open(pdf):
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
                out.append(len(ch))
    return out


def oxi_counts(docx, env_extra):
    env = dict(os.environ)
    env.update(env_extra)
    dump = os.path.join(OUT, "oxi.json")
    subprocess.run([GDI, docx, os.path.join(OUT, "oxi"), "96", "--dump-layout=" + dump],
                   capture_output=True, timeout=600, env=env)
    d = json.load(open(dump, encoding="utf-8"))
    out = []
    for pg in d["pages"]:
        rows = {}
        for el in pg.get("elements", []):
            if el.get("type") != "text":
                continue
            rows.setdefault(round(el.get("y", 0), 1), []).append(el)
        for y in sorted(rows):
            v = sorted(rows[y], key=lambda e: e["x"])
            t = "".join(e.get("text", "") for e in v)
            if t.startswith("甲"):
                out.append(len(t))
    return out


def main():
    docx, index = build()
    w = word_counts(docx)
    a = oxi_counts(docx, {"OXI_S1218_DISABLE": "1"})
    b = oxi_counts(docx, {"OXI_S1218": "1"})
    if not (len(w) == len(a) == len(b) == len(index)):
        print("counts: word %d oxi_off %d oxi_on %d arms %d" % (len(w), len(a), len(b), len(index)))
    by = {}
    for (name, r, n), cw, ca, cb in zip(index, w, a, b):
        by.setdefault(name, []).append((r / 20.0, cw, ca, cb, n))
    print("   arm        last r where line 1 still holds everything")
    print("   %-11s %8s %8s %8s" % ("", "Word", "S1218 off", "S1218 on"))
    for name, rows in by.items():
        n = rows[0][4]
        def flip(idx):
            full = [r for r, *rest in rows if rest[idx] >= n]
            return max(full) if full else None
        print("   %-11s %8s %8s %8s"
              % (name,
                 "%.2f" % flip(0) if flip(0) is not None else "-",
                 "%.2f" % flip(1) if flip(1) is not None else "-",
                 "%.2f" % flip(2) if flip(2) is not None else "-"))


if __name__ == "__main__":
    main()
