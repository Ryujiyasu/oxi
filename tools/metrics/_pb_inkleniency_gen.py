# -*- coding: utf-8 -*-
"""Pin the no-type-docGrid Latin auto-multiple page-bottom threshold (S1009 redo).

S1009 shipped "the flip is at the line's INK bottom" and implemented it as Oxi's
`ink_lh` (= the TYPO box, 1.0em for Calibri).  creative__006b6f8693b2c9d4
contradicts it: its 'Experience required for the role:' has ink bottom inside
the content bottom by 0.386pt yet Word pushes it.

Round 1 (4 arms x 10 T, Calibri 11 / line=259) showed every arm — a wrapped
continuation line, a <w:br/> continuation line, a fresh 1-line paragraph, and a
fresh paragraph after an 8pt gap — flipping at the SAME T, so the line CLASS is
not the discriminator.  The flip window (12.92, 13.67] contains neither the ink
box (11.0) nor the multiplied box (14.492) but does contain the NATURAL
(unmultiplied hhea) line, 1.2207 x 11 = 13.428.

Round 2 (this file) sweeps finer and adds a second font and two more multipliers
so "natural line" can be separated from any per-font constant.

Usage: python _pb_inkleniency_gen.py gen | measure | read
"""
import os
import sys
import json
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.abspath(os.path.join(HERE, "..", "..", "pipeline_data", "_pb_inkleniency"))

PAGE_H = 841.92           # A4
MARGIN = 72.0             # 1440tw top and bottom
CBOT = PAGE_H - MARGIN    # 769.92
GAP = 8.0                 # docDefaults after=160
MARKER_H = 12.0           # exact-height section marker line

# (id, font, half-points, w:line, natural-em)   natural-em = hhea (asc+desc+gap)/upm
CONFIGS = [
    ("C11a", "Calibri", 22, 259, 1.2207),          # the target document's regime
    ("T12b", "Times New Roman", 24, 360, 1.1499),  # other font, 1.5x
    ("C12c", "Calibri", 24, 480, 1.2207),          # 2.0x
]

# arms: W = wrapped continuation, N = fresh 1-line paragraph
ARMS = ("W", "N")

WRAP_TEXT = (
    "Delivery of the agreed territory plan requires consistent weekly contact "
    "with every account in the region together with accurate reporting of each "
    "visit so that the wider commercial team can respond quickly whenever a new "
    "opportunity appears anywhere inside the assigned area "
)

SECTPR = ('<w:pgSz w:w="11906" w:h="16838"/>'
          '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
          ' w:header="708" w:footer="708" w:gutter="0"/>'
          '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/>')


def cfg_geom(c):
    _id, font, hp, line, nat = c
    fs = hp / 2.0
    natural = nat * fs
    box = natural * (line / 240.0)
    ink = fs                      # Oxi's ink_lh for these fonts (typo box = 1.0em)
    return fs, natural, box, ink


def ts_for(c):
    """Sweep the target box-top across [box-flip - 1, ink-flip + 1] in 0.3pt steps."""
    _fs, natural, box, ink = cfg_geom(c)
    lo = CBOT - box - 1.0
    hi = CBOT - ink + 1.0
    n = int((hi - lo) / 0.3) + 1
    return [round(lo + 0.3 * i, 3) for i in range(n)]


def rpr(c):
    _id, font, hp, _line, _nat = c
    return (f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
            f'<w:sz w:val="{hp}"/>')


def exact_p(c, tw, text=""):
    r = rpr(c)
    body = (f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
            if text else "")
    return (f'<w:p><w:pPr><w:spacing w:after="0" w:line="{tw}" w:lineRule="exact"/>'
            f'<w:rPr>{r}</w:rPr></w:pPr>{body}</w:p>')


def filler(c, pt):
    tw = int(round(pt * 20))
    out = []
    while tw > 400:
        out.append(exact_p(c, 400))
        tw -= 400
    if tw > 0:
        out.append(exact_p(c, tw))
    return "".join(out)


def auto_p(c, text, after=0):
    r = rpr(c)
    line = c[3]
    return (f'<w:p><w:pPr><w:spacing w:after="{after}" w:line="{line}" w:lineRule="auto"/>'
            f'<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


def section(c, tag, arm, t):
    _fs, _natural, box, _ink = cfg_geom(c)
    offset = 2 * box if arm == "W" else box
    f_pt = t - MARGIN - MARKER_H - offset
    assert f_pt >= 0, (tag, f_pt)
    body = exact_p(c, int(round(MARKER_H * 20)), f"SEC{tag}") + filler(c, f_pt)
    if arm == "W":
        body += auto_p(c, WRAP_TEXT + f"TGT{tag}", after=0)
    else:
        body += auto_p(c, f"PRV{tag}", after=0) + auto_p(c, f"TGT{tag}", after=0)
    body += f'<w:p><w:pPr><w:sectPr>{SECTPR}</w:sectPr></w:pPr></w:p>'
    return body


def gen():
    os.makedirs(OUT, exist_ok=True)
    arms, body = [], ""
    for c in CONFIGS:
        for arm in ARMS:
            for i, t in enumerate(ts_for(c)):
                tag = f"{c[0]}{arm}{i:02d}"
                fs, natural, box, ink = cfg_geom(c)
                arms.append({"tag": tag, "cfg": c[0], "arm": arm, "t": t,
                             "natural": round(natural, 3), "box": round(box, 3),
                             "ink": round(ink, 3)})
                body += section(c, tag, arm, t)
    body += f'<w:sectPr>{SECTPR}</w:sectPr>'
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{body}</w:body></w:document>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
              '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr>'
              '<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>'
              '</w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style></w:styles>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '</Relationships>')
    path = os.path.join(OUT, "inkleniency.docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/_rels/document.xml.rels", drels)
    json.dump({"arms": arms, "cbot": CBOT}, open(os.path.join(OUT, "arms.json"), "w"), indent=1)
    print("gen", path, len(arms), "arms")


def measure():
    import win32com.client as win32
    src = os.path.join(OUT, "inkleniency.docx")
    pdf = os.path.join(OUT, "inkleniency.pdf")
    if os.path.exists(pdf):
        os.remove(pdf)
    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        d = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
        d.ExportAsFixedFormat(OutputFileName=pdf, ExportFormat=17)
        d.Close(False)
    finally:
        app.Quit()
    print("measured", pdf, os.path.getsize(pdf))


def read():
    import fitz
    meta = json.load(open(os.path.join(OUT, "arms.json")))
    d = fitz.open(os.path.join(OUT, "inkleniency.pdf"))
    import re as _re
    where = {}
    for pi in range(d.page_count):
        txt = d[pi].get_text()
        for m in _re.finditer(r"(SEC|TGT|PRV)([A-Za-z]\d{2}[a-z][WN]\d{2})", txt):
            where.setdefault(m.group(0), pi + 1)
    print("cbot=%.2f" % meta["cbot"])
    by = {}
    for a in meta["arms"]:
        by.setdefault((a["cfg"], a["arm"]), []).append(a)
    for k in sorted(by):
        rows = sorted(by[k], key=lambda a: a["t"])
        last_keep, first_push = None, None
        for a in rows:
            sp = where.get("SEC" + a["tag"])
            tp = where.get("TGT" + a["tag"])
            if sp is None or tp is None:
                continue
            if tp == sp:
                last_keep = a["t"]
            elif first_push is None and last_keep is not None:
                first_push = a["t"]
        a0 = rows[0]
        if last_keep is None or first_push is None:
            print("%-10s INCONCLUSIVE (last_keep=%s first_push=%s)" % (str(k), last_keep, first_push))
            continue
        lo = meta["cbot"] - first_push
        hi = meta["cbot"] - last_keep
        def mark(v):
            return "  <== " if lo < v <= hi else "      "
        print("%-10s threshold in (%.3f, %.3f]" % (str(k), lo, hi))
        print("            natural %7.3f%s" % (a0["natural"], mark(a0["natural"])))
        print("            box     %7.3f%s" % (a0["box"], mark(a0["box"])))
        print("            ink     %7.3f%s" % (a0["ink"], mark(a0["ink"])))
    d.close()


if __name__ == "__main__":
    {"gen": gen, "measure": measure, "read": read}[sys.argv[1]]()
