# -*- coding: utf-8 -*-
"""Where does 29dc6e8943fe's gridSpan=6 cell actually start?

Every paragraph in that cell sits 4.45pt to the right in Oxi, which costs the
③ line its last character under the derived-cell bundle. The drawn rules agree
(Word 54.48/159.98/174.14/437.62/537.10 against Oxi 54.45/159.95/174.10/437.2/
537.0), and the cell carries `tcBorders left=nil`, so its own left edge is never
drawn -- which is why reading it off the page has not settled anything.

So slice the row out and render it BOTH ways from the same file.

    python _pb_29dc_cell.py
"""
import glob
import os
import re
import shutil
import subprocess
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_29dc_cell")
SRC = [p for p in glob.glob(os.path.join(REPO, "tools", "golden-test", "documents",
                                         "docx", "29dc6e*.docx"))
       if not os.path.basename(p).startswith("~$")][0]
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
MARK = "授業科目の目的"


ARMS = [
    ("base", lambda h, r: (h, r)),
    ("no_gridspan", lambda h, r: (h, re.sub(r'<w:gridSpan w:val="\d+"/>', "", r))),
    ("no_tcborders", lambda h, r: (h, re.sub(r"<w:tcBorders>.*?</w:tcBorders>", "", r, flags=re.S))),
    ("style_a", lambda h, r: (h, r.replace('<w:pStyle w:val="ac"/>', '<w:pStyle w:val="a"/>'))),
    ("no_track", lambda h, r: (h, re.sub(r'<w:spacing w:val="-9"/>', "", r))),
    ("no_lineexact", lambda h, r: (h, re.sub(r'<w:spacing w:line="259" w:lineRule="exact"/>', "", r))),
    ("no_left_tw", lambda h, r: (h, r.replace('<w:ind w:left="81" w:firstLineChars="100" w:firstLine="199"/>',
                                              '<w:ind w:firstLineChars="100" w:firstLine="199"/>'))),
    ("no_ind", lambda h, r: (h, r.replace('<w:ind w:left="81" w:firstLineChars="100" w:firstLine="199"/>', ''))),
    # read the cell margin directly: no ind at all, and the margin set explicitly
    ("mar0", lambda h, r: (h.replace("<w:tblLayout", '<w:tblCellMar><w:left w:w="0" w:type="dxa"/>'
                                     '<w:right w:w="0" w:type="dxa"/></w:tblCellMar><w:tblLayout'),
                           r.replace('<w:ind w:left="81" w:firstLineChars="100" w:firstLine="199"/>', ''))),
    ("mar108", lambda h, r: (h.replace("<w:tblLayout", '<w:tblCellMar><w:left w:w="108" w:type="dxa"/>'
                                       '<w:right w:w="108" w:type="dxa"/></w:tblCellMar><w:tblLayout'),
                             r.replace('<w:ind w:left="81" w:firstLineChars="100" w:firstLine="199"/>', ''))),
    ("leftchars_too", lambda h, r: (h, r.replace('<w:ind w:left="81" w:firstLineChars="100" w:firstLine="199"/>',
                                                 '<w:ind w:leftChars="50" w:left="81" w:firstLineChars="100" w:firstLine="199"/>'))),
]


def build_arms():
    """One file per ablation; read where Word puts the ③ line each time."""
    os.makedirs(OUT, exist_ok=True)
    x = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    i = x.index(MARK)
    tbl_start = x.rindex("<w:tbl>", 0, i)
    tbl_head = x[tbl_start:x.index("</w:tblGrid>", tbl_start) + len("</w:tblGrid>")]
    rs = max((x.rfind("<w:tr ", 0, i), x.rfind("<w:tr>", 0, i)))
    row = x[rs:x.index("</w:tr>", i) + len("</w:tr>")]
    head = x[:x.index("<w:body>") + len("<w:body>")]
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", x, re.S).group(0)
    sect = re.sub(r"<w:(headerReference|footerReference)[^>]*/>", "", sect)
    out = []
    for name, fn in ARMS:
        h, r = fn(tbl_head, row)
        doc = head + h + r + "</w:tbl>" + sect + "</w:body></w:document>"
        dst = os.path.join(OUT, "arm_%s.docx" % name)
        shutil.copyfile(SRC, dst)
        zin = zipfile.ZipFile(SRC)
        zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = doc.encode("utf-8")
            zout.writestr(item, data)
        zout.close()
        out.append((name, dst))
    return out


def arms_main():
    import win32com.client as wc
    import fitz
    files = build_arms()
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    rows = []
    try:
        for name, docx in files:
            pdf = os.path.splitext(docx)[0] + ".pdf"
            d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf),
                                  ExportFormat=17, OpenAfterExport=False)
            d.Close(False)
            doc = fitz.open(pdf)
            rules, x0 = [], None
            for page in doc:
                for dr in page.get_drawings():
                    for it in dr["items"]:
                        if it[0] == "l" and abs(it[1].x - it[2].x) < 0.4 and abs(it[1].y - it[2].y) > 3:
                            rules.append(round((it[1].x + it[2].x) / 2, 2))
                        elif it[0] == "re" and it[1].width < 0.9 and it[1].height > 3:
                            rules.append(round(it[1].x0, 2))
                for b in page.get_text("rawdict")["blocks"]:
                    if b["type"] != 0:
                        continue
                    for l in b["lines"]:
                        ch = sorted([c for s in l["spans"] for c in s["chars"]],
                                    key=lambda c: c["origin"][0])
                        if ch and "".join(c["c"] for c in ch).startswith("③") and x0 is None:
                            x0 = ch[0]["origin"][0]
            rows.append((name, sorted(set(rules)), x0))
    finally:
        app.Quit()
    print("   %-14s %-10s %-9s %s" % ("arm", "cell left", "text x", "indent from the cell edge"))
    for name, rules, x0 in rows:
        cell = [r for r in rules if r > 150][0] if any(r > 150 for r in rules) else 0
        print("   %-14s %-10.2f %-9s %s"
              % (name, cell, "%.2f" % x0 if x0 else "-",
                 "%+.2f" % (x0 - cell - 5.4) if x0 else "-"))


def build():
    os.makedirs(OUT, exist_ok=True)
    x = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    i = x.index(MARK)
    tbl_start = x.rindex("<w:tbl>", 0, i)
    tbl_head = x[tbl_start:x.index("</w:tblGrid>", tbl_start) + len("</w:tblGrid>")]
    rs = max((x.rfind("<w:tr ", 0, i), x.rfind("<w:tr>", 0, i)))
    row = x[rs:x.index("</w:tr>", i) + len("</w:tr>")]
    head = x[:x.index("<w:body>") + len("<w:body>")]
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", x, re.S).group(0)
    sect = re.sub(r"<w:(headerReference|footerReference)[^>]*/>", "", sect)
    print("row: %d cells, gridSpans %s, tcW %s"
          % (row.count("<w:tc>"), re.findall(r'<w:gridSpan w:val="(\d+)"', row),
             re.findall(r'<w:tcW w:w="(\d+)"', row)))
    print("grid cols:", re.findall(r'w:w="(\d+)"', tbl_head[tbl_head.index("<w:tblGrid>"):]))
    doc = head + tbl_head + row + "</w:tbl>" + sect + "</w:body></w:document>"
    dst = os.path.join(OUT, "row.docx")
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = doc.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    return dst


def word_pdf(docx):
    import win32com.client as wc
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
    return pdf


def word_read(pdf):
    import fitz
    doc = fitz.open(pdf)
    rules, lines = [], []
    for page in doc:
        for d in page.get_drawings():
            for it in d["items"]:
                if it[0] == "l" and abs(it[1].x - it[2].x) < 0.4 and abs(it[1].y - it[2].y) > 3:
                    rules.append(round((it[1].x + it[2].x) / 2, 2))
                elif it[0] == "re" and it[1].width < 0.9 and it[1].height > 3:
                    rules.append(round(it[1].x0, 2))
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                ch = sorted([c for s in l["spans"] for c in s["chars"]],
                            key=lambda c: c["origin"][0])
                if ch:
                    lines.append(("".join(c["c"] for c in ch)[:14], ch[0]["origin"][0]))
    return sorted(set(rules)), lines


def oxi_read(docx):
    import json
    dump = os.path.join(OUT, "row_oxi.json")
    subprocess.run([GDI, docx, os.path.join(OUT, "row_oxi"), "96",
                    "--dump-layout=" + dump], capture_output=True, timeout=300)
    d = json.load(open(dump, encoding="utf-8"))
    rules, lines = [], []
    for pg in d["pages"]:
        for el in pg.get("elements", []):
            if el.get("type") == "border" and (el.get("w") or 0) < 0.9 and (el.get("h") or 0) > 3:
                rules.append(round(el["x"], 2))
            elif el.get("type") == "text":
                lines.append((el.get("text", "")[:14], el["x"], round(el.get("y", 0), 1)))
    return sorted(set(rules)), lines


def main():
    docx = build()
    wr, wl = word_read(word_pdf(docx))
    orr, ol = oxi_read(docx)
    print("\nWord rules:", wr[:8])
    print("Oxi  rules:", orr[:8])
    # group Oxi's text by line, keep the leftmost x
    seen = {}
    for t, x, y in ol:
        seen.setdefault(y, (t, x))
        if x < seen[y][1]:
            seen[y] = (t, x)
    ow = {t: x for t, x in seen.values()}
    print("\n   %-16s %8s %8s %8s" % ("line", "Word x", "Oxi x", "delta"))
    for t, x in wl[:18]:
        key = next((k for k in ow if k[:6] == t[:6]), None)
        if key is None:
            continue
        print("   %-16s %8.2f %8.2f %+8.2f" % (t[:14], x, ow[key], ow[key] - x))


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "arms":
        arms_main()
    else:
        main()
