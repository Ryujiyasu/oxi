# -*- coding: utf-8 -*-
"""How far does Word COMPRESS inter-word spaces to keep one more word on a line?

`reference__0042471c` p1 showed Word packing a token Oxi wraps, by squeezing the
line's 15 spaces from their natural 2.193pt to 1.683pt (0.767x) to absorb a
7.96pt overflow.  Over that whole document the extreme was 0.713x and 25% of
justified lines sat below natural.  That is one specimen, so it is a hypothesis
until a controlled probe says what the RULE is.

Design -- one arm per page, so an arm is identified by its page index and no
marker text perturbs the widths:

    a justified paragraph of N identical words, in a column whose width is
    swept in fine steps by a right indent.

As the column narrows, the last word of line 1 eventually drops to line 2.  The
step just BEFORE the drop is the maximum compression Word will accept, and the
space advance measured there is the limit.  Sweeping the WORD LENGTH varies the
number of spaces per line, which separates the two candidate models:

    ratio model   space >= r_min x natural        -> absorbable overflow scales
                                                     with the space COUNT
    budget model  overflow <= a fixed slack       -> absorbable overflow is flat

  python _pb_spacecomp_gen.py gen
  python _pb_spacecomp_gen.py read      # Word truth (export + per-line spans)
  python _pb_spacecomp_gen.py oxi       # Oxi, same arms
"""
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_spacecomp")
DOCX = os.path.join(OUT, "spacecomp.docx")   # rebound in __main__ per mode
PDF = os.path.join(OUT, "spacecomp_word.pdf")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FONT, SZ = "Cambria", 20                # 10pt
SPACE_PT = 0.2202 * 10.0                # Cambria hmtx space @10pt = 2.2021
# Word lengths -> different space COUNT per line (the model discriminator)
WORD_LENS = [2, 4, 8]
# right indent sweep, in twips: 0 .. 600tw (0 .. 30pt) in 20tw (1pt) steps
INDENTS = list(range(0, 1620, 20))

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"')
# ★compatibilityMode is THE discriminator for this whole question (2026-08-14):
# 15 lets Word over-fill a justified line and compress the spaces, <=14 is greedy.
# A probe with NO settings.xml falls back to the old mode and measures the WRONG
# Word. Selected with `mode=<n>` on the command line; default 15.
MODE = 15
def SETTINGS():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + '>'
            '<w:compat><w:compatSetting w:name="compatibilityMode"'
            ' w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat>'
            '</w:settings>' % MODE)
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/>'
          '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '</w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/></w:style></w:styles>' % (FONT, FONT, FONT, SZ, SZ))


def arms():
    """justified 群のあと、同一内容・同一段幅の LEFT 揃え双子群。

    左揃え行の語数 = その段幅に NATURAL で入る語数。justified 行がそれより
    多ければ、その差がまさに Word の over-fill であり、語幅を justify 結果から
    逆算する必要が無くなる（語幅のデバイス量子化ノイズが消える）。"""
    base = [(wl, ind) for wl in WORD_LENS for ind in INDENTS]
    return [(wl, ind, "both") for wl, ind in base] + [(wl, ind, "left") for wl, ind in base]


def body():
    ps = []
    for i, (wl, ind, jc) in enumerate(arms()):
        # enough words to fill 3+ lines at the widest column
        words = " ".join(["m" * wl] * 60)
        ppr = ('<w:pPr><w:jc w:val="%s"/>' % jc
               + ('<w:pageBreakBefore/>' if i else '')
               + '<w:ind w:right="%d"/>' % ind
               + '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
               '</w:pPr>')
        ps.append('<w:p>%s<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ppr, words))
    return "".join(ps)


def gen():
    os.makedirs(OUT, exist_ok=True)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '>'
           '<w:body>' + body() +
           '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
           ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS())
        z.writestr("word/document.xml", doc)
    print("wrote %s  %d arms (column 468pt minus 0..30pt indent)" % (DOCX, len(arms())))


def _report(rows, who):
    """rows[i] = (n_words_line1, n_spaces_line1, median_space_adv, line1_width)"""
    print("%s   natural space = %.4f pt" % (who, SPACE_PT))
    print("%-5s %-6s %8s %8s %9s %9s %9s  %s"
          % ("wlen", "indent", "col_pt", "words", "spaces", "space", "ratio", "note"))
    prev_key = None
    for (wl, ind, jc), r in zip(arms(), rows):
        if r is None:
            print("%-5d %-6d   MISSING" % (wl, ind))
            continue
        if jc != "both":
            continue
        nw, ns, adv, w1 = r
        col = 468.0 - ind / 20.0
        drop = ""
        if prev_key is not None and prev_key[0] == wl and nw < prev_key[1]:
            drop = "  <-- word DROPPED here"
        print("%-5d %-6d %8.2f %8d %9d %9.4f %9.4f%s"
              % (wl, ind, col, nw, ns, adv, adv / SPACE_PT, drop))
        prev_key = (wl, nw)


def word_pdf():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(DOCX, ReadOnly=True)
    try:
        d.ExportAsFixedFormat(PDF, 17)
    finally:
        d.Close(False)
        app.Quit()


def _pdf_rows():
    import fitz
    from collections import defaultdict
    import statistics
    doc = fitz.open(PDF)
    rows = []
    for pi in range(doc.page_count):
        d = doc.load_page(pi).get_text("rawdict")
        lines = defaultdict(list)
        for blk in d["blocks"]:
            if blk.get("type") != 0:
                continue
            for ln in blk.get("lines", []):
                for sp in ln.get("spans", []):
                    for c in sp.get("chars", []):
                        lines[round(c["origin"][1], 0)].append((c["origin"][0], c["c"], c["bbox"][2]))
        if not lines:
            rows.append(None)
            continue
        y0 = sorted(lines)[0]
        cs = sorted(lines[y0])
        adv = [cs[i + 1][0] - cs[i][0] for i in range(len(cs) - 1) if cs[i][1] == " "]
        txt = "".join(c[1] for c in cs)
        rows.append((len(txt.split()), len(adv),
                     statistics.median(adv) if adv else 0.0,
                     max(c[2] for c in cs) - cs[0][0]))
    doc.close()
    return rows


def read():
    if not os.path.exists(PDF) or "--refresh" in sys.argv:
        word_pdf()
    _report(_pdf_rows(), "WORD")


def oxi(envs=""):
    import json
    import statistics
    from collections import defaultdict
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    dump = os.path.join(OUT, "spacecomp_oxi.json")
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
        # element-to-element advance across the gap = word advance + space
        gaps = []
        for a, b in zip(es, es[1:]):
            g = b["x"] - (a["x"] + a.get("w", 0.0))
            if g > 0.1:
                gaps.append(g)
        nw = len(es)
        rows.append((nw, len(gaps),
                     statistics.median(gaps) if gaps else 0.0,
                     max(e["x"] + e.get("w", 0.0) for e in es) - es[0]["x"]))
    _report(rows, "OXI  " + (envs or "(default)"))


if __name__ == "__main__":
    for a in sys.argv[1:]:
        if a.startswith("mode="):
            MODE = int(a.split("=", 1)[1])
        if a.startswith("sz="):          # half-points, e.g. sz=40 -> 20pt
            SZ = int(a.split("=", 1)[1])
            SPACE_PT = 0.2202 * (SZ / 2.0)
            STYLES = STYLES.replace('w:val="20"', 'w:val="%d"' % SZ)
    DOCX = os.path.join(OUT, "spacecomp_m%d_sz%d.docx" % (MODE, SZ))
    PDF = os.path.join(OUT, "spacecomp_m%d_sz%d_word.pdf" % (MODE, SZ))
    cmd = sys.argv[1]
    if cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "read": read}[cmd]()
