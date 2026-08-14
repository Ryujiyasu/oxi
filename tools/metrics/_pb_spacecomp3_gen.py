# -*- coding: utf-8 -*-
"""Does the overflow CAP depend on the ADDED word's width, or on the space count?

`_pb_spacecomp_gen.py` filled the whole line with identical words, so the added
word's width and the line's space count always moved TOGETHER (longer words ->
fewer spaces).  Its floor+cap model (floor 0.1683em, cap ~0.61em) fit wlen
2/4/8 with zero errors but broke on wlen=1 (13/81), and the confound makes the
cause undecidable in that harness.

This probe decouples them: every line starts with the same run of single-'m'
words (fixing the space count), and only the LAST word's length varies.

    m m m ... m  WWWW      <- prefix fixed, W swept 1/2/4/8/12 chars

If the cap is a property of the LINE (total absorbable overflow), the
transition columns for different W lengths differ by exactly the width
difference of W, and the absorbed amount at the transition is CONSTANT.
If it depends on the added word itself, the absorbed amount will vary with W.

Both a justified arm and a LEFT twin are generated per (W, indent) so the
over-fill is read directly (justified word count minus left word count),
never inferred from widths.

  python _pb_spacecomp3_gen.py gen
  python _pb_spacecomp3_gen.py read      # Word truth via PDF
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_spacecomp")
DOCX = os.path.join(OUT, "spacecomp3.docx")   # rebound in __main__ per size
PDF = os.path.join(OUT, "spacecomp3_word.pdf")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FONT, SZ = "Cambria", 20                 # half-points; sz=40 -> 20pt via argv
NPREFIX = 20                              # 'm' words before the candidate
WLENS = [1, 2, 4, 8, 12]
INDENTS = list(range(0, 820, 20))         # 0..41pt in 1pt steps

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
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
    base = [(wl, ind) for wl in WLENS for ind in INDENTS]
    return [(wl, ind, "both") for wl, ind in base] + [(wl, ind, "left") for wl, ind in base]


def gen():
    os.makedirs(OUT, exist_ok=True)
    ps = []
    for i, (wl, ind, jc) in enumerate(arms()):
        # prefix of single-m words, then the candidate, then enough text that
        # the paragraph has 2+ lines (so line 1 is justified), repeated tail
        text = " ".join(["m"] * NPREFIX) + " " + "m" * wl + " " + " ".join(["m"] * 30)
        ppr = ('<w:pPr><w:jc w:val="%s"/>' % jc
               + ('<w:pageBreakBefore/>' if i else '')
               + '<w:ind w:right="%d"/>' % ind
               + '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>')
        ps.append('<w:p>%s<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ppr, text))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '>'
           '<w:body>' + "".join(ps) +
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
    print("wrote %s  %d arms (prefix %d x 'm', W in %s)"
          % (DOCX, len(arms()), NPREFIX, WLENS))


def read():
    import fitz
    import statistics
    from collections import defaultdict
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
    doc = fitz.open(PDF)
    if doc.page_count != len(arms()):
        raise SystemExit("PAGE/ARM MISMATCH: %d pages vs %d arms — a paragraph "
                         "spilled; every downstream number would be read off the "
                         "wrong arm (the 687-vs-648 lesson)." % (doc.page_count, len(arms())))

    def line1(pi):
        d = doc.load_page(pi).get_text("rawdict")
        rows = defaultdict(list)
        for blk in d["blocks"]:
            if blk.get("type") != 0:
                continue
            for ln in blk.get("lines", []):
                for sp in ln.get("spans", []):
                    for c in sp.get("chars", []):
                        rows[round(c["origin"][1], 0)].append((c["origin"][0], c["c"]))
        return sorted(rows[sorted(rows)[0]]) if rows else None

    half = len(arms()) // 2
    NAT = 2.161 * (SZ / 20.0)
    print("wlen  col      just語 left語 over  space(比率)   吸収pt")
    prev_key = None
    for i in range(half):
        wl, ind, _ = arms()[i]
        J = line1(i)
        L = line1(half + i)
        if not J or not L:
            continue
        nwj = len("".join(c[1] for c in J).split())
        nwl = len("".join(c[1] for c in L).split())
        adv = [J[k + 1][0] - J[k][0] for k in range(len(J) - 1) if J[k][1] == " "]
        sp = statistics.median(adv) if adv else 0.0
        col = (12240 - 2880 - ind) / 20.0
        over = nwj - nwl
        absorbed = (nwj - 1) * (NAT - sp) if over > 0 else 0.0
        mark = ""
        if prev_key is not None and prev_key[0] == wl and over < prev_key[1]:
            mark = "  <-- DROP"
        print("%-4d %8.2f %6d %6d %5d  %7.4f(%5.3f) %8.2f%s"
              % (wl, col, nwj, nwl, over, sp, sp / NAT, absorbed, mark))
        prev_key = (wl, over)
    doc.close()


if __name__ == "__main__":
    for a in sys.argv[2:]:
        if a.startswith("sz="):
            SZ = int(a.split("=", 1)[1])
    DOCX = os.path.join(OUT, "spacecomp3_sz%d.docx" % SZ)
    PDF = os.path.join(OUT, "spacecomp3_sz%d_word.pdf" % SZ)
    STYLES = STYLES.replace('w:val="20"', 'w:val="%d"' % SZ)
    {"gen": gen, "read": read}[sys.argv[1]]()
