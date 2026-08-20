# -*- coding: utf-8 -*-
"""Page-bottom fit under a no-type docGrid: does Word gate the push on the
EXACT cumulative or the px-ROUNDED position?

The last underived cell of the S1181 quantizer (see _pb_gridwalk_gen.py and
the 2026-08-20 ra_archive notes): placement is round_px(exact absolute), but
nyserda/uklocal/legal_0001482d/002c1ffa A-B flips all sit at page-bottom (or
table row-fit) knife edges, so the FIT coordinate is unknown.

Each arm is its own page: [pbb marker][47 filler lines @276][target before=X].
The target's landing page (1 = fit, 2 = pushed) is a robust binary.  Two
target heights bracket the models on both sides:
  - txt240 (18.398px): capacity-minus-line frac ~.13 -> the ROUNDED model
    keeps fitting ~0.37px (7.4tw) PAST the exact threshold
  - txt480 (36.797px): frac ~.73 -> the ROUNDED model stops fitting ~0.23px
    BEFORE the exact threshold (sign flip)
A model must match both transition positions AND signs.

    python _pb_gridfit_gen.py gen
    python _pb_gridfit_gen.py read
    python _pb_gridfit_gen.py oxi [ENV=..]
"""
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_gridfit")
DOCX = os.path.join(OUT, "gridfit.docx")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

FONT, SZ = "Arial", 24

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
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
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/>'
          '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>'
          '</w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/>'
          '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '</w:style></w:styles>')

N_FILL = 45
# (label, target_line, before_tw sweep)
GROUPS = [
    ("F240", 240, list(range(554, 575, 2))),
    ("F480", 480, list(range(278, 299, 2))),
]


def rpr():
    return ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (FONT, FONT, FONT, SZ, SZ))


def txt(line, text, before=None, pbb=False):
    sp = ('<w:spacing %sw:line="%d" w:lineRule="auto"/>'
          % (('w:before="%d" ' % before) if before is not None else "", line))
    return ('<w:p><w:pPr>%s<w:widowControl w:val="0"/>%s%s</w:pPr>'
            '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % ("<w:pageBreakBefore/>" if pbb else "", sp, rpr(), rpr(), text))


def arms():
    out = []
    for glbl, line, sweep in GROUPS:
        for x in sweep:
            out.append((glbl, line, x))
    return out


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (glbl, line, x) in enumerate(arms()):
        body.append(txt(240, "M%02d %s x=%d" % (ai, glbl, x), pbb=True))
        for i in range(N_FILL):
            body.append(txt(276, "fill %d" % i))
        body.append(txt(line, "TGT%02d" % ai, before=x))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/>'
           '<w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(arms()), "arms x", N_FILL + 2, "paras")


def report(rows, who):
    # rows: {arm_index: (marker_page, target_page, target_y)}
    print(who)
    NAT = 13.7988
    for ai, (glbl, line, x) in enumerate(arms()):
        mp, tp, ty = rows.get(ai, (None, None, None))
        fit = "?" if tp is None else ("FIT" if tp == mp else "PUSH")
        # model predictions: stream = 18.3984 + 47*21.15816 + before_px
        stream = 48.0 + NAT / 0.75 + N_FILL * (NAT * 1.15) / 0.75 + (x / 20.0) / 0.75
        lh = NAT * (line / 240.0) / 0.75
        cap = (841.89 - 72.0) / 0.75 + 48.0  # content bottom abs px
        e_fit = stream + lh <= cap + 1e-6
        r_fit = round(stream) + lh <= cap + 1e-6
        print("%s x=%3d  %-4s  y=%s   exact:%s rounded:%s%s"
              % (glbl, x, fit, "-" if ty is None else "%.2f" % ty,
                 "FIT " if e_fit else "PUSH", "FIT " if r_fit else "PUSH",
                 "   <-- discriminates" if e_fit != r_fit else ""))


def read():
    import re
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.ScreenUpdating = False
    d = app.Documents.Open(DOCX, ReadOnly=True)
    rows = {}
    try:
        d.Repaginate()
        mp = {}
        for p in d.Paragraphs:
            rng = p.Range
            t = rng.Text
            c = d.Range(rng.Start, rng.Start)
            m = re.match(r"M(\d\d) ", t)
            if m:
                mp[int(m.group(1))] = c.Information(3)
            m = re.match(r"TGT(\d\d)", t)
            if m:
                ai = int(m.group(1))
                rows[ai] = (mp.get(ai), c.Information(3), round(c.Information(6), 2))
    finally:
        d.Close(False)
        app.Quit()
    report(rows, "WORD")


def oxi(envs=""):
    import json
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "gridfit_oxi.json")
    subprocess.run([GDI, DOCX, os.path.join(tempfile.gettempdir(), "gridfit"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    rows = {}
    mp = {}
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        for e in pg["elements"]:
            t = (e.get("text") or "").strip()
            if t.startswith("M") and len(t) > 3 and t[1:3].isdigit():
                mp[int(t[1:3])] = pg["page"]
            if t.startswith("TGT"):
                ai = int(t[3:5])
                rows[ai] = (mp.get(ai), pg["page"], e["y"])
    report(rows, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "read": read}[sys.argv[1]]()
