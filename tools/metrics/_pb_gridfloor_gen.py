# -*- coding: utf-8 -*-
"""Where does a JUSTIFIED line end when the grid does not divide the measure?

`_pb_gridpitch_gen.py` showed the BREAK happens as if the line were only
floor(content/pitch) * pitch wide -- with charSpace=1966 that hides 7.96pt of a
425.2pt measure. This asks the other half of the question: does Word also END
the justified line there (the text area really is narrower), or does it stretch
the line out to the right margin anyway (the truncation is a break-time rule)?

Each arm is a jc=both paragraph long enough to wrap, at one charSpace. Read the
right edge of every line that is NOT the paragraph's last.

    python _pb_gridfloor_gen.py            # build, export, measure
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_gridfloor")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
LEFT_PT = 1701 / 20.0                      # 85.05
CONTENT_PT = (11906 - 1701 - 1701) / 20.0  # 425.2
DEFAULT_FS = 10.5
CHAR_SPACES = [1966, 1453, 532, 0]
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')


def build_one(cs):
    paras = []
    for fs in (10.5, 9.0):
        txt = "甲" + "亜" * 200
        paras.append(
            '<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="both"/>'
            '<w:ind w:left="0" w:right="0"/>'
            '<w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (int(fs * 2), int(fs * 2), txt))
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    grid = ('<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="%d"/>' % cs
            if cs else '<w:docGrid w:type="lines" w:linePitch="360"/>')
    sect = re.sub(r"<w:docGrid[^>]*/>", grid, sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s><w:body>%s%s</w:body></w:document>'
           % (W_NS, "".join(paras), sect))
    dst = os.path.join(OUT, "floor%d.docx" % cs)
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    return dst


def main():
    import fitz
    import win32com.client as wc
    os.makedirs(OUT, exist_ok=True)
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    try:
        for cs in CHAR_SPACES:
            docx = build_one(cs)
            pdf = os.path.join(OUT, "floor%d.pdf" % cs)
            d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf),
                                  ExportFormat=17, OpenAfterExport=False)
            d.Close(False)
    finally:
        app.Quit()
    for cs in CHAR_SPACES:
        cs_pt = cs / 4096.0
        pitch = DEFAULT_FS + cs_pt
        cells = int(CONTENT_PT / pitch)
        floored = cells * pitch if cs else CONTENT_PT
        doc = fitz.open(os.path.join(OUT, "floor%d.pdf" % cs))
        ends = {}
        for page in doc:
            for b in page.get_text("rawdict")["blocks"]:
                if b["type"] != 0:
                    continue
                for l in b["lines"]:
                    ch = sorted([c for s in l["spans"] for c in s["chars"]],
                                key=lambda c: c["origin"][0])
                    if len(ch) < 10:
                        continue
                    size = round(l["spans"][0]["size"], 2)
                    # the line's right edge = last origin + one em
                    em = ch[-1]["bbox"][2] - ch[-1]["bbox"][0]
                    ends.setdefault(size, []).append(ch[-1]["origin"][0] + em)
        print(chr(10) + "charSpace=%d  pitch=%.4f  cells=%d  floored=%.2f (margin right = %.2f)"
              % (cs, pitch, cells, floored, LEFT_PT + CONTENT_PT))
        for size in sorted(ends):
            v = sorted(ends[size])
            mid = v[len(v) // 2]
            print("   pdf size %5.2f  n=%3d  median right edge %7.2f   "
                  "full margin %7.2f (%+.2f)   floored %7.2f (%+.2f)"
                  % (size, len(v), mid, LEFT_PT + CONTENT_PT,
                     mid - (LEFT_PT + CONTENT_PT), LEFT_PT + floored,
                     mid - (LEFT_PT + floored)))


if __name__ == "__main__":
    main()
