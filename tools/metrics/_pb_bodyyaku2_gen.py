# -*- coding: utf-8 -*-
"""Where does tokyoshugyo's ⑧ line get 0.74em of credit from?

_pb_bodyyaku_gen.py measured the plain body line: half an em, flat, whatever the
class and however many marks -- the cell-side law holds here too. But the real
line spends 0.74em (two 、 compressed 3.9pt each). The two things that line has
and the probe does not are a run of FULLWIDTH SPACES and a numbering MARKER with
a hanging indent. Sweep both.

    python _pb_bodyyaku2_gen.py gen
    python _pb_bodyyaku2_gen.py pdf
"""
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_bodyyaku_gen as B  # noqa: E402

OUT = os.path.join(B.REPO, "pipeline_data", "_pb_bodyyaku2")
EM, NCH = B.EM, B.NCH
R_TW = list(range(0, 401, 5))     # 0..20pt in 0.25pt steps

MARKER = ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="37"/></w:numPr>'
          '<w:ind w:leftChars="0" w:left="884" w:hanging="425"/>')


def body_chars(n_mark, n_space, mark="、"):
    """NCH-1 characters after 甲: n_mark marks up front, n_space spaces after."""
    body = ["亜"] * (NCH - 1)
    for i in range(n_mark):
        body[3 + i * 6] = mark
    # the spaces sit as ONE run late in the line, exactly as the real line has it
    start = NCH - 8
    for i in range(n_space):
        body[start + i] = "　"
    return "甲" + "".join(body)


def arms():
    out = []
    for j in range(0, 5):                       # spaces only
        out.append(("SP%d" % j, body_chars(0, j), ""))
    for j in range(0, 5):                       # spaces + two marks
        out.append(("S2%d" % j, body_chars(2, j), ""))
    for j in (0, 4):                            # + numbering marker
        out.append(("M2%d" % j, body_chars(2, j), MARKER))
    out.append(("U24", body_chars(2, 4), "", True))   # spaces underlined
    return out


def build():
    os.makedirs(OUT, exist_ok=True)
    paras, index = [], []
    for arm in arms():
        name, text, extra = arm[0], arm[1], arm[2]
        underline = len(arm) > 3
        for r in R_TW:
            index.append((name, r))
            if underline:
                head, tail = text.split("　", 1)
                nsp = text.count("　")
                tail = tail[nsp - 1:]
                runs = ('<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                        '<w:t xml:space="preserve">%s</w:t></w:r>'
                        '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/>'
                        '<w:u w:val="single"/></w:rPr>'
                        '<w:t xml:space="preserve">%s</w:t></w:r>'
                        '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                        '<w:t xml:space="preserve">%s</w:t></w:r>'
                        % (head, "　" * nsp, tail))
            else:
                runs = ('<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                        '<w:t xml:space="preserve">%s</w:t></w:r>' % text)
            ind = ('<w:ind w:left="0" w:right="%d"/>' % r if not extra
                   else '<w:ind w:leftChars="0" w:left="884" w:hanging="425" '
                        'w:right="%d"/>' % r)
            ppr = ('<w:pPr><w:pStyle w:val="a"/>'
                   + (extra.split("<w:ind")[0] if extra else "")
                   + ind + '</w:pPr>')
            paras.append("<w:p>" + ppr + runs + "</w:p>")
    src = zipfile.ZipFile(B.SRC)
    doc = src.read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s><w:body>%s%s</w:body></w:document>'
           % (B.W_NS, "".join(paras), sect))
    dst = os.path.join(OUT, "bodyyaku2.docx")
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in src.infolist():
        data = src.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    with open(os.path.join(OUT, "arms.txt"), "w", encoding="utf-8") as fh:
        for a in index:
            fh.write("%s\t%d\n" % a)
    print("built %s  (%d paragraphs, %d arms)" % (dst, len(paras), len(arms())))


def to_pdf():
    import win32com.client as wc
    docx = os.path.join(OUT, "bodyyaku2.docx")
    pdf = os.path.join(OUT, "bodyyaku2.pdf")
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
    index = [l.split("\t") for l in open(os.path.join(OUT, "arms.txt"),
             encoding="utf-8").read().splitlines()]
    doc = fitz.open(os.path.join(OUT, "bodyyaku2.pdf"))
    lines = []
    for page in doc:
        rows = []
        for blk in page.get_text("rawdict").get("blocks", []):
            for ln in blk.get("lines", []):
                cs = [c for sp in ln["spans"] for c in sp.get("chars", [])]
                t = "".join(c["c"] for c in cs).rstrip()
                if t:
                    rows.append((round(ln["bbox"][1], 1), t))
        rows.sort()
        lines.extend(rows)
    paras, cur = [], None
    for y, t in lines:
        if "甲" in t[:3]:
            if cur:
                paras.append(cur)
            cur = [t]
        elif cur is not None:
            cur.append(t)
    if cur:
        paras.append(cur)
    print("arms %d, paragraphs %d" % (len(index), len(paras)))
    res = {}
    for (name, r), p in zip(index, paras):
        res.setdefault(name, []).append((int(r), len(p)))
    base = 425.2 - NCH * EM
    print("\narm   keep<=r      split>=r   credit pt    em")
    for name in sorted(res):
        rows = sorted(res[name])
        one = [r for r, k in rows if k == 1]
        spl = [r for r, k in rows if k > 1]
        if not one:
            print("%-5s never one line" % name)
            continue
        rmax, rmin = max(one), (min(spl) if spl else None)
        credit = rmax / 20.0 - base
        flag = "" if rmin == rmax + 5 else "  (NON-MONOTONE)"
        print("%-5s %4d (%5.2fpt) %s  %7.3f  %.4f em%s"
              % (name, rmax, rmax / 20.0,
                 ("%4d" % rmin) if rmin is not None else "   -",
                 credit, credit / EM, flag))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "gen":
        build()
    elif cmd == "pdf":
        to_pdf(); measure()
    else:
        measure()
