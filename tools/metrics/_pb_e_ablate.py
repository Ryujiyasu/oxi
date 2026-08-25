# -*- coding: utf-8 -*-
"""Slice tokyoshugyo's （エ） paragraph and knock features out one at a time.

Word wraps its final 「と。」 although every synthetic axis says a line-final 。
is free. The slice must reproduce the wrap; then one ablation should flip it.

    python _pb_e_ablate.py
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
OUT = os.path.join(REPO, "pipeline_data", "_pb_e_ablate")
SRC = [p for p in glob.glob(os.path.join(REPO, "tools", "golden-test", "documents",
                                         "docx", "tokyoshugyo*.docx"))
       if "~$" not in os.path.basename(p)][0]
MARK = "自己申告した労働時間を超えて"

ARMS = [
    ("base", lambda p: p),
    ("no_numpr", lambda p: re.sub(r"<w:numPr>.*?</w:numPr>|<w:tabs>.*?</w:tabs>", "", p, flags=re.S)),
    ("no_track", lambda p: re.sub(r'<w:spacing w:val="-9"/>', "", p)),
    ("no_leftchars", lambda p: p.replace(' w:leftChars="0"', "")),
    ("one_run", None),          # handled specially
    ("ta_for_to", lambda p: p.replace("こと。</w:t>", "こ亜。</w:t>")),
    ("no_hanging", lambda p: re.sub(r'<w:ind[^>]*/>', '<w:ind w:left="884"/>', p, count=1)),
]


def one_run(p):
    """Merge every run into one, keeping the FIRST run's rPr."""
    ppr = re.search(r"<w:pPr>.*?</w:pPr>", p, re.S).group(0)
    texts = "".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", p, re.S))
    head = p[:p.index(">") + 1]
    first_r = re.search(r"<w:r(?:\s[^>]*)?>(?:(?!</w:r>).)*?</w:r>", p, re.S).group(0)
    m = re.search(r"<w:rPr>.*?</w:rPr>", first_r, re.S)
    return (head + ppr + "<w:r>" + (m.group(0) if m else "")
            + '<w:t xml:space="preserve">' + texts + "</w:t></w:r></w:p>")


def build():
    os.makedirs(OUT, exist_ok=True)
    x = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    i = x.index(MARK)
    ps = max(x.rfind("<w:p ", 0, i), x.rfind("<w:p>", 0, i))
    pe = x.index("</w:p>", i) + len("</w:p>")
    para = x[ps:pe]
    head = x[:x.index("<w:body>") + len("<w:body>")]
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", x, re.S).group(0)
    sect = re.sub(r"<w:(headerReference|footerReference)[^>]*/>", "", sect)
    out = []
    for name, fn in ARMS:
        p2 = one_run(para) if name == "one_run" else fn(para)
        doc = head + p2 + sect + "</w:body></w:document>"
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


def main():
    import win32com.client as wc
    import fitz
    files = build()
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
            lines = []
            for b in doc[0].get_text("rawdict")["blocks"]:
                if b["type"] != 0:
                    continue
                for l in b["lines"]:
                    ch = sorted([c for s in l["spans"] for c in s["chars"]],
                                key=lambda c: c["origin"][0])
                    if ch:
                        lines.append((round(l["bbox"][1], 1),
                                      "".join(c["c"] for c in ch).strip()))
            lines = [t for _, t in sorted(lines, key=lambda x: x[0]) if t]
            rows.append((name, lines))
    finally:
        app.Quit()
    print("original Word: line 2 ends 「…確認するこ」, line 3 = 「と。」(wraps)")
    for name, lines in rows:
        tails = " | ".join(t[-6:] for t in lines)
        verdict = "WRAPS (=original)" if any(t in ("と。", "亜。") for t in lines) \
            else "hangs (flipped!)"
        print("   %-12s %d lines: %s   -> %s" % (name, len(lines), tails, verdict))


if __name__ == "__main__":
    main()
