# -*- coding: utf-8 -*-
"""Read the BUDGET of d77a58's ヨ/カ cell instead of inferring it.

Word breaks 「…適用されることもありま | す。」 there; Oxi with OXI_YAKUCOMP keeps す。
on line 1. Every quantity around it agrees to 0.1pt (glyph advances, cell rules,
pool=0 under S1207), so the decision hangs on the cell's text budget -- which
cannot be read off a justified line without assuming what is being tested.

So take the row out of the document verbatim, replicate it with the paragraph's
RIGHT INDENT swept in 0.25pt steps, and read the width at which Word's own break
moves. Arm r=0 must reproduce the original break, or the slice is not faithful.

    python _pb_d77a_budget.py            # build, export, report
"""
import glob
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_d77a_budget")
SRC = [p for p in glob.glob(os.path.join(REPO, "tools", "golden-test", "documents",
                                         "docx", "d77a*.docx"))
       if not os.path.basename(p).startswith("~$")][0]
MARK = "ウェブサイト全体"
# ★The table is `tblW auto` with no tblLayout, so Word AUTOFITS the column: a
# right indent on the paragraph just widens the column and the text keeps its
# width (the first sweep moved nothing at all, 81 arms identical). So sweep the
# COLUMN itself under a fixed layout instead. The grid value the full document
# renders is 8952tw; the arm at 8952 must reproduce Word's own break.
GRID_TW = 8952
W_TW = list(range(GRID_TW - 200, GRID_TW + 60, 5))   # -10pt..+3pt in 0.25pt steps
BORDERS = [4, 8, 16, 24]                             # eighths of a point
W_BORDER = list(range(GRID_TW - 20, GRID_TW + 121, 5))


def slice_row():
    x = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    i = x.index(MARK)
    tbl_start = x.rindex("<w:tbl>", 0, i)
    tbl_head = x[tbl_start:x.index("</w:tblGrid>", tbl_start) + len("</w:tblGrid>")]
    rs = max((x.rfind("<w:tr ", 0, i), x.rfind("<w:tr>", 0, i)))
    row = x[rs:x.index("</w:tr>", i) + len("</w:tr>")]
    head = x[:x.index("<w:body>") + len("<w:body>")]
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", x, re.S).group(0)
    sect = re.sub(r"<w:(headerReference|footerReference)[^>]*/>", "", sect)
    return head, tbl_head, row, sect


def fixed_at(tbl_head, row, w_tw):
    """The same row in a FIXED-layout table of the given column width."""
    head = tbl_head.replace('<w:tblW w:w="0" w:type="auto"/>',
                            '<w:tblW w:w="%d" w:type="dxa"/>'
                            '<w:tblLayout w:type="fixed"/>' % w_tw)
    head = re.sub(r'<w:gridCol w:w="[0-9]+"/>', '<w:gridCol w:w="%d"/>' % w_tw, head)
    row = re.sub(r'<w:tcW w:w="[0-9]+" w:type="dxa"/>',
                 '<w:tcW w:w="%d" w:type="dxa"/>' % w_tw, row, count=1)
    return head, row



def with_border(head, row, sz):
    """Set every rule of the table AND the cell to `sz` eighths of a point."""
    def bump(seg):
        return re.sub(r'(<w:(?:top|left|bottom|right|insideH|insideV)[^>]*?)'
                      r'w:sz="[0-9]+"', r'\g<1>w:sz="%d"' % sz, seg)
    head = re.sub(r'<w:tblBorders>.*?</w:tblBorders>',
                  lambda m: bump(m.group(0)), head, flags=re.S)
    row = re.sub(r'<w:tcBorders>.*?</w:tcBorders>',
                 lambda m: bump(m.group(0)), row, flags=re.S)
    return head, row


def build_border():
    """One table per (border size, column width): does the text area inset by
    half the rule? Declared before measuring -- if it does, the width at which
    the break moves grows by (sz - 4)/8 pt; if the rule is ignored, it stays."""
    os.makedirs(OUT, exist_ok=True)
    head0, tbl_head, row0, sect = slice_row()
    blocks, arms = [], []
    for sz in BORDERS:
        for w in W_BORDER:
            h, rr = fixed_at(tbl_head, row0, w)
            h, rr = with_border(h, rr, sz)
            arms.append((sz, w))
            blocks.append(h + rr + "</w:tbl>"
                          + '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>')
    doc = head0 + "".join(blocks) + sect + "</w:body></w:document>"
    dst = os.path.join(OUT, "d77a_border.docx")
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = doc.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    with open(os.path.join(OUT, "border_arms.txt"), "w", encoding="utf-8") as fh:
        for a in arms:
            fh.write("%d" % a[0] + chr(9) + "%d" % a[1] + chr(10))
    print("built %s (%d arms)" % (dst, len(arms)))


def measure_border():
    import fitz
    arms = [tuple(int(v) for v in l.split(chr(9))) for l in
            open(os.path.join(OUT, "border_arms.txt"), encoding="utf-8").read().splitlines()]
    doc = fitz.open(os.path.join(OUT, "d77a_border.pdf"))
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
            t = "".join(c["c"] for c in ch)
            # ★the row holds TWO paragraphs that open with 本利用ルール -- ヨ
            # (…適用されることが…) and カ (…適用されることも…). Matching the
            # prefix alone returned 156 lines for 116 arms.
            if t.startswith("本利用ルール") and "ことも" in t:
                heads.append(len(ch))
    if len(heads) != len(arms):
        print("%d marked lines for %d arms" % (len(heads), len(arms)))
        return
    by = {}
    for (sz, w), n in zip(arms, heads):
        by.setdefault(sz, []).append((w, n))
    print("   rule    width where す。 first fits    predicted if the text area")
    print("   (pt)                                  insets by half the rule")
    base = None
    for sz in BORDERS:
        got = [w for w, n in by[sz] if n >= 43]
        first = min(got) / 20.0 if got else None
        if base is None and first is not None:
            base = first
        pred = base + (sz - BORDERS[0]) / 8.0 if base is not None else None
        print("   %5.3f   %s                        %s"
              % (sz / 8.0,
                 "%8.2f" % first if first else "  not in window",
                 "%8.2f" % pred if pred else "-"))


def build():
    os.makedirs(OUT, exist_ok=True)
    head, tbl_head, row, sect = slice_row()
    blocks = []
    for w in W_TW:
        h, rr = fixed_at(tbl_head, row, w)
        blocks.append(h + rr + "</w:tbl>"
                      + '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>')
    doc = head + "".join(blocks) + sect + "</w:body></w:document>"
    dst = os.path.join(OUT, "d77a_budget.docx")
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = doc.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    print("built %s (%d arms, %d..%dtw)"
          % (dst, len(W_TW), W_TW[0], W_TW[-1]))


def to_pdf():
    import win32com.client as wc
    docx = os.path.join(OUT, "d77a_budget.docx")
    pdf = os.path.join(OUT, "d77a_budget.pdf")
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
    doc = fitz.open(os.path.join(OUT, "d77a_budget.pdf"))
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
            t = "".join(c["c"] for c in ch)
            if t.startswith("本利用ルール") and "ことも" in t:
                heads.append((len(ch), ch[-1]["origin"][0], ch[0]["origin"][0], t[-3:]))
    if len(heads) != len(W_TW):
        print("%d marked lines for %d arms -- check the slice" % (len(heads), len(W_TW)))
    print("   col(pt)  d(pt)  glyphs   line x0    last origin  tail")
    prev = None
    for (n, lastx, x0, tail), w in zip(heads, W_TW):
        flip = "  <<< break moves" if prev is not None and n != prev else ""
        print("   %7.2f %+6.2f   %3d    %8.2f   %8.2f     %s%s"
              % (w / 20.0, (w - GRID_TW) / 20.0, n, x0, lastx, tail, flip))
        prev = n


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "all"
    if cmd == "border":
        build_border()
        import win32com.client as wc
        app = wc.Dispatch("Word.Application")
        app.Visible = False
        try:
            d = app.Documents.Open(os.path.abspath(
                os.path.join(OUT, "d77a_border.docx")), ReadOnly=True)
            d.ExportAsFixedFormat(OutputFileName=os.path.abspath(
                os.path.join(OUT, "d77a_border.pdf")), ExportFormat=17,
                OpenAfterExport=False)
            d.Close(False)
        finally:
            app.Quit()
        measure_border()
        sys.exit()
    if cmd in ("gen", "all"):
        build()
    if cmd in ("pdf", "all"):
        to_pdf()
    measure()
