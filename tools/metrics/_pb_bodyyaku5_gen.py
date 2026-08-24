# -*- coding: utf-8 -*-
"""How far does the trailing 　 unlock the 約物 pool?

Measured so far (all with a numbering marker, ＭＳ 明朝 10.5, jc=both, compat 11):
  marks only, any count            -> 0.5em
  fullwidth spaces only            -> 0
  2 marks + >=1 space next to the
  line's last character            -> 1.0em, and the render spends 0.5em on EACH 、
  1 mark  + spaces next to it      -> 0.5em
  2 marks + spaces six chars back  -> 0.5em

So the trailing space releases a second half em. This arm set asks how many it
releases (n marks x one trailing space), whether the releasing character has to be
a space or any compressible, and how close to the end it has to sit.

    python _pb_bodyyaku5_gen.py gen
    python _pb_bodyyaku5_gen.py pdf
"""
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_bodyyaku_gen as B  # noqa: E402

OUT = os.path.join(B.REPO, "pipeline_data", "_pb_bodyyaku5")
R_TW = list(range(0, 601, 5))
NCH = 36
MARKER = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="37"/></w:numPr>'
IND_M = 'w:leftChars="0" w:left="884" w:hanging="425"'
MARK_POS = [6, 12, 18, 24]          # head positions for up to four marks


def text_of(nmark, fills):
    """36 characters: 火 + 亜 filler, nmark marks in the head, `fills` = {idx: ch}."""
    t = ["火"] + ["亜"] * (NCH - 2) + ["に"]
    for i in range(nmark):
        t[MARK_POS[i]] = "、"
    for idx, ch in fills.items():
        t[idx] = ch
    return "".join(t)


ARMS = []
for n in range(0, 5):
    ARMS.append(("n%d_bare" % n, text_of(n, {})))
    ARMS.append(("n%d_sp34" % n, text_of(n, {34: "　"})))
ARMS += [
    ("n2_sp33", text_of(2, {33: "　"})),
    ("n2_sp31", text_of(2, {31: "　"})),
    ("n2_sp28", text_of(2, {28: "　"})),
    ("n2_yk34", text_of(2, {34: "、"})),
    ("n2_pd34", text_of(2, {34: "。"})),
    ("n2_cl34", text_of(2, {34: "）"})),
    ("n2_sp3234", text_of(2, {32: "　", 34: "　"})),
    ("n3_sp34b", text_of(3, {34: "　"})),
]


def build():
    os.makedirs(OUT, exist_ok=True)
    paras, index = [], []
    for name, txt in ARMS:
        assert len(txt) == NCH, (name, len(txt))
        for r in R_TW:
            index.append((name, r))
            paras.append(
                '<w:p><w:pPr><w:pStyle w:val="a7"/>' + MARKER
                + '<w:ind %s w:right="%d"/></w:pPr>' % (IND_M, r)
                + '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                  '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % txt)
    src = zipfile.ZipFile(B.SRC)
    doc = src.read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s><w:body>%s%s</w:body></w:document>'
           % (B.W_NS, "".join(paras), sect))
    dst = os.path.join(OUT, "bodyyaku5.docx")
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in src.infolist():
        data = src.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    open(os.path.join(OUT, "arms.txt"), "w", encoding="utf-8").write(
        "".join("%s\t%d\n" % a for a in index))
    print("built %s (%d paragraphs, %d arms)" % (dst, len(paras), len(ARMS)))


def to_pdf():
    import win32com.client as wc
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(os.path.join(OUT, "bodyyaku5.docx"), ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=os.path.join(OUT, "bodyyaku5.pdf"),
                              ExportFormat=17, OpenAfterExport=False)
        d.Close(False)
    finally:
        app.Quit()


def rows():
    import fitz
    index = [l.split("\t") for l in open(os.path.join(OUT, "arms.txt"),
             encoding="utf-8").read().splitlines()]
    doc = fitz.open(os.path.join(OUT, "bodyyaku5.pdf"))
    lines = []
    for page in doc:
        rs = []
        for blk in page.get_text("rawdict").get("blocks", []):
            for ln in blk.get("lines", []):
                cs = [c for sp in ln["spans"] for c in sp.get("chars", [])]
                t = "".join(c["c"] for c in cs).rstrip()
                if t:
                    rs.append((round(ln["bbox"][1], 1), t, cs))
        rs.sort(); lines.extend(rs)
    paras, cur = [], None
    for y, t, cs in lines:
        if "火亜" in t:
            if cur:
                paras.append(cur)
            cur = [(t, cs)]
        elif cur is not None:
            cur.append((t, cs))
    paras.append(cur)
    return index, paras


def measure():
    index, paras = rows()
    print("arms %d paragraphs %d" % (len(index), len(paras)))
    if len(index) != len(paras):
        print("!! grouping mismatch"); return
    res = {}
    for (name, r), p in zip(index, paras):
        res.setdefault(name, []).append((int(r), len(p), p))
    keep = {}
    for name, _ in ARMS:
        rr = sorted(res[name])
        one = [r for r, k, _ in rr if k == 1]
        spl = [r for r, k, _ in rr if k > 1]
        keep[name] = max(one) if one else None
        mono = "" if (one and spl and min(spl) == max(one) + 5) else "  (NON-MONOTONE)"
        print("%-11s keep<=%s split>=%s%s" % (name,
              ("%4d (%5.2fpt)" % (keep[name], keep[name] / 20.0)) if one else "  none    ",
              ("%4d" % min(spl)) if spl else "  -", mono))
    z = keep.get("n0_bare")
    print("\ncredit against n0_bare (%s tw)" % z)
    for name, _ in ARMS:
        if keep.get(name) is None or z is None:
            continue
        d = (keep[name] - z) / 20.0
        print("  %-11s %6.3f pt  %.4f em" % (name, d, d / B.EM))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "gen":
        build()
    elif cmd == "pdf":
        to_pdf(); measure()
    else:
        measure()
