# -*- coding: utf-8 -*-
"""Does Word keep an OVER-FULL justified line, or wrap it?  (specimen repro)

`_pb_spacecomp_gen.py` (identical repeated words) says Word is strictly GREEDY:
the moment the natural line exceeds the column it drops the last word, even when
keeping it would cost under 1% of space compression.  Yet
`reference__0042471c` has 50 justified lines whose spaces sit at 0.72x natural,
which can only happen if the line holds more than fits.

Both cannot be right, and the arithmetic that called those lines "over-full by
8pt" is not trustworthy: Word device-rounds every advance (its natural Cambria
space measures 2.161 = 18px at 600dpi, not the hmtx 2.193), so a 110-char sum
carries several pt of slop -- the same order as the overflow being claimed.

So: take the specimen's own line VERBATIM, put it in a column of the specimen's
own width, and let Word answer directly.  Arms sweep the column so the same text
is asked at, above and below its natural width.

  python _pb_spacecomp2_gen.py gen
  python _pb_spacecomp2_gen.py read
  python _pb_spacecomp2_gen.py oxi
"""
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_spacecomp2")
DOCX = os.path.join(OUT, "spacecomp2.docx")
PDF = os.path.join(OUT, "spacecomp2_word.pdf")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FONT, SZ = "Cambria", 20
# The specimen's p4 line, verbatim, plus the words that follow it on the next
# line -- so Word has a real choice about where to break.
TEXT = ("usual least squares method, which was used in Eviews 9.0 software. "
        "Time series data models are used to forecast the future using "
        "historical data from 2015 to 2020. In essence, regression analysis "
        "is used to estimate or predict the value of the dependent variable.")
# The specimen's column is 470.18pt.  Sweep +-6pt around it in 0.5pt steps.
INDENTS = [i for i in range(-120, 130, 10)]      # twips, 0.5pt steps

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/>'
          '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '</w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/></w:style></w:styles>' % (FONT, FONT, FONT, SZ, SZ))


def arms():
    return list(INDENTS)


def gen():
    os.makedirs(OUT, exist_ok=True)
    ps = []
    for i, ind in enumerate(arms()):
        ppr = ('<w:pPr><w:jc w:val="both"/>'
               + ('<w:pageBreakBefore/>' if i else '')
               + '<w:ind w:right="%d"/>' % ind
               + '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>')
        ps.append('<w:p>%s<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ppr, TEXT))
    # left 85.1pt = 1702tw, right margin set so the column is 470.18 at indent 0
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '>'
           '<w:body>' + "".join(ps) +
           '<w:sectPr><w:pgSz w:w="12240" w:h="20160"/>'
           '<w:pgMar w:top="1440" w:right="1134" w:bottom="1440" w:left="1702"'
           ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    print("wrote %s   %d arms, column = %.2f .. %.2f pt"
          % (DOCX, len(arms()), (12240 - 1702 - 1134 - max(INDENTS)) / 20.0,
             (12240 - 1702 - 1134 - min(INDENTS)) / 20.0))


def _rows_from_pdf():
    import fitz
    import statistics
    from collections import defaultdict
    doc = fitz.open(PDF)
    out = []
    for pi in range(doc.page_count):
        d = doc.load_page(pi).get_text("rawdict")
        lines = defaultdict(list)
        for blk in d["blocks"]:
            if blk.get("type") != 0:
                continue
            for ln in blk.get("lines", []):
                for sp in ln.get("spans", []):
                    for c in sp.get("chars", []):
                        lines[round(c["origin"][1], 0)].append((c["origin"][0], c["c"]))
        if not lines:
            out.append(None)
            continue
        cs = sorted(lines[sorted(lines)[0]])
        adv = [cs[i + 1][0] - cs[i][0] for i in range(len(cs) - 1) if cs[i][1] == " "]
        txt = "".join(c[1] for c in cs).rstrip()
        out.append((txt, statistics.median(adv) if adv else 0.0, len(adv)))
    doc.close()
    return out


def _report(rows, who, natural):
    print("%s   (natural space measured on the widest arm's last line)" % who)
    print("%-8s %8s %7s %9s %8s  %s" % ("indent", "col_pt", "words", "space", "ratio", "line 1 ends"))
    for ind, r in zip(arms(), rows):
        if r is None:
            print("%-8d   MISSING" % ind)
            continue
        txt, adv, ns = r
        col = (12240 - 1702 - 1134 - ind) / 20.0
        print("%-8d %8.2f %7d %9.4f %8.4f  ...%r"
              % (ind, col, len(txt.split()), adv, adv / natural if natural else 0, txt[-30:]))


def read():
    if not os.path.exists(PDF) or "--refresh" in sys.argv:
        import win32com.client as w
        app = w.DispatchEx("Word.Application")
        app.Visible = False
        d = app.Documents.Open(DOCX, ReadOnly=True)
        try:
            d.ExportAsFixedFormat(PDF, 17)
        finally:
            d.Close(False)
            app.Quit()
    _report(_rows_from_pdf(), "WORD", 2.161)


def oxi(envs=""):
    import json
    import statistics
    from collections import defaultdict
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    dump = os.path.join(OUT, "spacecomp2_oxi.json")
    subprocess.run([GDI, DOCX, os.path.join(OUT, "px"), "--dump-layout=" + dump],
                   check=True, capture_output=True, env=env)
    d = json.load(open(dump, encoding="utf-8"))
    rows = []
    for pg in d["pages"]:
        lines = defaultdict(list)
        for e in pg["elements"]:
            if (e.get("text") or "").strip():
                lines[round(e["y"], 1)].append(e)
        if not lines:
            rows.append(None)
            continue
        es = sorted(lines[sorted(lines)[0]], key=lambda e: e["x"])
        gaps = [b["x"] - (a["x"] + a.get("w", 0.0)) for a, b in zip(es, es[1:])]
        gaps = [g for g in gaps if g > 0.1]
        rows.append(("".join(e["text"] for e in es), statistics.median(gaps) if gaps else 0.0, len(gaps)))
    _report(rows, "OXI  " + (envs or "(default)"), 2.161)


if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "read": read}[cmd]()
