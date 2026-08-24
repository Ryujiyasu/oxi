# -*- coding: utf-8 -*-
"""Replica sweep of tokyoshugyo's ⑧ line, with each suspect knocked out in turn.

The plain-line probe says half an em, flat. The real line spends 0.74em. So take
the REAL paragraph and remove one feature at a time -- the marks, the fullwidth
spaces, the underline, the numbering marker -- sweeping the right indent for each,
and read the credit off the arm that loses it.

Every arm's geometry is calibrated by its OWN control (the same shape with the
marks replaced by ordinary characters), so the indent the marker really applies
never has to be guessed.

    python _pb_bodyyaku3_gen.py gen
    python _pb_bodyyaku3_gen.py pdf
"""
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_bodyyaku_gen as B  # noqa: E402

OUT = os.path.join(B.REPO, "pipeline_data", "_pb_bodyyaku3")
R_TW = list(range(0, 401, 5))          # 0..20pt in 0.25pt steps
HEAD = "火災等非常災害の発生を発見したときは、直ちに臨機の措置をとり、"
SP = "　　　　"
TAIL = "に"
MARKER = ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="37"/></w:numPr>')
IND_M = 'w:leftChars="0" w:left="884" w:hanging="425"'


def variant(marks, spaces):
    h = HEAD if marks else HEAD.replace("、", "亜")
    s = SP if spaces else "亜" * len(SP)
    return h, s, TAIL


ARMS = [
    # name, marker?, marks?, spaces?, underline?
    ("R_all", True, True, True, True),
    ("R_noU", True, True, True, False),
    ("R_noS", True, True, False, False),
    ("R_noM", True, False, True, False),
    ("R_none", True, False, False, False),
    ("P_all", False, True, True, False),
    ("P_noM", False, False, True, False),
    ("P_noS", False, True, False, False),
    ("P_none", False, False, False, False),
]


def build():
    os.makedirs(OUT, exist_ok=True)
    paras, index = [], []
    for name, marker, marks, spaces, under in ARMS:
        h, s, t = variant(marks, spaces)
        for r in R_TW:
            index.append((name, r))
            if under:
                runs = ('<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                        '<w:t xml:space="preserve">%s</w:t></w:r>'
                        '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/>'
                        '<w:u w:val="single"/></w:rPr>'
                        '<w:t xml:space="preserve">%s</w:t></w:r>'
                        '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                        '<w:t xml:space="preserve">%s</w:t></w:r>' % (h, s, t))
            else:
                runs = ('<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                        '<w:t xml:space="preserve">%s</w:t></w:r>' % (h + s + t))
            if marker:
                ppr = ('<w:pPr><w:pStyle w:val="a7"/>' + MARKER
                       + '<w:ind %s w:right="%d"/></w:pPr>' % (IND_M, r))
            else:
                ppr = ('<w:pPr><w:pStyle w:val="a"/>'
                       '<w:ind w:left="0" w:right="%d"/></w:pPr>' % r)
            paras.append("<w:p>" + ppr + runs + "</w:p>")
    src = zipfile.ZipFile(B.SRC)
    doc = src.read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s><w:body>%s%s</w:body></w:document>'
           % (B.W_NS, "".join(paras), sect))
    dst = os.path.join(OUT, "bodyyaku3.docx")
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in src.infolist():
        data = src.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    open(os.path.join(OUT, "arms.txt"), "w", encoding="utf-8").write(
        "".join("%s\t%d\n" % a for a in index))
    print("built %s  (%d paragraphs, %d arms)" % (dst, len(paras), len(ARMS)))


def to_pdf():
    import win32com.client as wc
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(os.path.join(OUT, "bodyyaku3.docx"), ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=os.path.join(OUT, "bodyyaku3.pdf"),
                              ExportFormat=17, OpenAfterExport=False)
        d.Close(False)
    finally:
        app.Quit()


def measure():
    import fitz
    index = [l.split("\t") for l in open(os.path.join(OUT, "arms.txt"),
             encoding="utf-8").read().splitlines()]
    doc = fitz.open(os.path.join(OUT, "bodyyaku3.pdf"))
    lines = []
    for page in doc:
        rows = []
        for blk in page.get_text("rawdict").get("blocks", []):
            for ln in blk.get("lines", []):
                cs = [c for sp in ln["spans"] for c in sp.get("chars", [])]
                t = "".join(c["c"] for c in cs).rstrip()
                if t:
                    rows.append((round(ln["bbox"][1], 1), t, cs))
        rows.sort()
        lines.extend(rows)
    paras, cur = [], None
    for y, t, cs in lines:
        if "火災等" in t:
            if cur:
                paras.append(cur)
            cur = [(t, cs)]
        elif cur is not None:
            cur.append((t, cs))
    if cur:
        paras.append(cur)
    print("arms %d, paragraphs %d" % (len(index), len(paras)))
    if len(index) != len(paras):
        print("!! grouping mismatch")
        return
    res = {}
    for (name, r), p in zip(index, paras):
        res.setdefault(name, []).append((int(r), len(p), p))
    print("\narm      keep<=r    split>=r   (a control's keep is its zero)")
    keep = {}
    for name, _, _, _, _ in ARMS:
        rows = sorted(res[name])
        one = [r for r, k, _ in rows if k == 1]
        spl = [r for r, k, _ in rows if k > 1]
        keep[name] = max(one) if one else None
        print("%-8s %s  %s%s"
              % (name, ("%4d (%5.2fpt)" % (keep[name], keep[name] / 20.0))
                 if one else "   none    ",
                 ("%4d" % min(spl)) if spl else "   -",
                 "" if (spl and one and min(spl) == max(one) + 5) else "  (NON-MONOTONE)"))
    print("\ncredit vs its own control (pt / em)")
    for arm, ctrl in (("R_all", "R_none"), ("R_noU", "R_none"),
                      ("R_noS", "R_none"), ("R_noM", "R_none"),
                      ("P_all", "P_none"), ("P_noM", "P_none"),
                      ("P_noS", "P_none")):
        if keep.get(arm) is None or keep.get(ctrl) is None:
            continue
        d = (keep[arm] - keep[ctrl]) / 20.0
        print("  %-8s - %-8s = %6.3f pt   %.4f em" % (arm, ctrl, d, d / B.EM))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "gen":
        build()
    elif cmd == "pdf":
        to_pdf(); measure()
    else:
        measure()
