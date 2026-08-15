# -*- coding: utf-8 -*-
"""How much does a TALL footer take off the body?

technical__002c1ffa (0.8493, pcd -1 over 368 pages) has a footer of 8 paragraphs
plus a table under `pgMar bottom=4252tw (212.6pt)` / `footer=3402tw (170.1pt)`.
Word's last body line on its TOC pages sits at y=573 (+line = 583); Oxi keeps
filling to 596 (+line = 606) — about 20pt too low — and that extra line per page
walks 641 paragraphs off by one.

The footer lives inside the bottom margin. Word only pushes the body up when the
footer is TALLER than the room the footer distance leaves it, and the question is
what exactly it counts: the footer's own height, its height plus the distance, a
minimum gap to the body, or the declared bottom margin as a floor.

Each arm is one page whose body is a run of numbered lines (so the last one that
fits is readable) under a footer of a swept size: N plain paragraphs, and a
variant with a table of R rows. Reading the last body line's y from the Word PDF
gives the body bottom directly.

  python _pb_ftrheight_gen.py gen
  python _pb_ftrheight_gen.py pdf      # Word truth
  python _pb_ftrheight_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_ftrheight")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

# The specimen's own geometry: A4, bottom margin 212.6pt, footer distance 170.1pt
PGW, PGH = 11907, 16839
MAR_TOP, MAR_BOT, MAR_LR = 1418, 4252, 2410
FTR_DIST = 3402
BODY_LINES = 60                      # enough to overflow every arm

# (name, footer paragraphs, footer table rows)
ARMS = [
    ("f1_p1_t0", 1, 0),
    ("f2_p2_t0", 2, 0),
    ("f3_p4_t0", 4, 0),
    ("f4_p8_t0", 8, 0),
    ("f5_p12_t0", 12, 0),
    ("f6_p1_t1", 1, 1),
    ("f7_p1_t3", 1, 3),
    ("f8_p8_t1", 8, 1),
    # ★the specimen's OWN footer part, verbatim: the synthetic arms agree with
    # Word to 0.4pt, so whatever costs technical__002c1ffa its 3pt lives in that
    # footer (8 paragraphs, three of them `w:spacing w:before="120"`, a 1-row
    # table, and a trailing `Footer`-styled empty paragraph).
    ("f9_real", -1, 0),
    # ★the real footer's paragraphs carry `w:spacing w:before="120"` (6pt) and
    # sz=16 (8pt). These arms isolate the space-before: same paragraph count as
    # f4, once with the 6pt before and once with the 8pt font.
    ("f10_p8_sb", -2, 0),
    ("f11_p8_sz16", -3, 0),
    # ★which declaration form does the estimate drop? f10 (direct before) is
    # dropped; creative__00d0925f's footer (direct LINE only, spacing from a
    # style) regressed when S1131 folded it, so that form is kept. These arms
    # separate the four combinations so the rule is measured, not inferred.
    ("f12_style_sb_directline", -4, 0),   # style before + direct line=auto
    ("f13_style_sb_only", -5, 0),         # style before, no direct spacing
    ("f14_direct_sb_exact", -6, 0),       # direct before + direct line=exact
    ("f15_style_sb_exact", -7, 0),        # style before + direct line=exact
    # ★creative__00d0925f's footer carries space-AFTER (style after=180) and no
    # before, and it PASSES without any fold — so after may not extend the
    # footer the way before does. These two arms test the asymmetry directly.
    ("f16_style_sa", -8, 0),              # style after + direct line=auto
    ("f17_style_sb_sa", -9, 0),           # style before AND after
    # ★the specimen's footer shape, rebuilt synthetically so one piece can be
    # removed at a time: [empty 8pt para with before=120] + [2-row table whose
    # ROW 1 cells carry before=120] + [trailing empty para]. f9_real leaves an
    # 11.1pt residual once its DOCPROPERTY fields resolve; these locate it.
    ("r1_replica", -10, 0),
    ("r2_no_rowsb", -11, 0),     # same, row-1 cells without before
    ("r3_no_tail", -12, 0),      # same as r1, no trailing empty para
    ("r4_no_head", -13, 0),      # same as r1, no leading empty para
]
SPECIMEN = os.path.join(REPO, "pipeline_data", "docx_corpus", "en", "technical",
                        "002c1ffa65f3a566.docx")


def docx():
    return os.path.join(OUT, "ftrheight.docx")


def rpr():
    return ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="20"/></w:rPr>')


def para(text, pbb=False):
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/>%s</w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t>'
            "</w:r></w:p>" % ("<w:pageBreakBefore/>" if pbb else "", rpr(), rpr(), text))


def footer_xml(npara, nrows):
    if npara == -1:                    # the specimen's own footer part
        return zipfile.ZipFile(SPECIMEN).read("word/footer4.xml").decode("utf-8")
    if npara in (-10, -11, -12, -13):
        def p8(txt, sb):
            return ('<w:p><w:pPr>%s<w:rPr><w:rFonts w:ascii="Times New Roman" '
                    'w:hAnsi="Times New Roman"/><w:sz w:val="16"/></w:rPr></w:pPr>'
                    '%s</w:p>'
                    % ('<w:spacing w:before="120"/>' if sb else "",
                       ('<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" '
                        'w:hAnsi="Times New Roman"/><w:sz w:val="16"/></w:rPr>'
                        '<w:t>%s</w:t></w:r>' % txt) if txt else ""))
        # ★2 rows made the stack too short: the body bottom stayed pinned to the
        # declared margin in all four arms and nothing discriminated. Six rows
        # push the stack past (bottom margin - footer distance) = 42.5pt so a
        # per-row spacing error shows up as whole lines.
        row_sb = npara != -11
        rows = ""
        for r in range(6):
            cells = "".join(
                '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>%s</w:tc>'
                % p8("R%dC%d" % (r + 1, c + 1), row_sb and r % 2 == 1) for c in range(3))
            rows += "<w:tr>%s</w:tr>" % cells
        tbl = ('<w:tbl><w:tblPr><w:tblW w:w="7303" w:type="dxa"/></w:tblPr>'
               '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/>'
               '<w:gridCol w:w="2503"/></w:tblGrid>' + rows + "</w:tbl>")
        head = "" if npara == -13 else p8("", True)
        tail = "" if npara == -12 else p8("", False)
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + NS + ">"
                + head + tbl + tail + "</w:ftr>")
    if npara in (-8, -9):
        style = "FtrSA" if npara == -8 else "FtrSBSA"
        body = "".join(
            '<w:p><w:pPr><w:pStyle w:val="%s"/>'
            '<w:spacing w:line="240" w:lineRule="auto"/>'
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="20"/></w:rPr></w:pPr><w:r>'
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="20"/></w:rPr><w:t>F%d</w:t></w:r></w:p>' % (style, i + 1)
            for i in range(8))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + NS + ">"
                + body + "</w:ftr>")
    if npara in (-4, -5, -6, -7):
        style = "FtrSB"
        direct = {-4: '<w:spacing w:line="240" w:lineRule="auto"/>',
                  -5: "",
                  -6: '<w:spacing w:before="120" w:line="200" w:lineRule="exact"/>',
                  -7: '<w:spacing w:line="200" w:lineRule="exact"/>'}[npara]
        use_style = npara in (-4, -5, -7)
        body = "".join(
            '<w:p><w:pPr>%s%s<w:rPr><w:rFonts w:ascii="Times New Roman" '
            'w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr></w:pPr><w:r>'
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="20"/></w:rPr><w:t>F%d</w:t></w:r></w:p>'
            % ('<w:pStyle w:val="%s"/>' % style if use_style else "", direct, i + 1)
            for i in range(8))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + NS + ">"
                + body + "</w:ftr>")
    if npara in (-2, -3):
        sb = ' w:before="120"' if npara == -2 else ""
        sz = 16 if npara == -3 else 20
        body = "".join(
            '<w:p><w:pPr><w:spacing%s w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="%d"/></w:rPr></w:pPr><w:r>'
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="%d"/></w:rPr><w:t>F%d</w:t></w:r></w:p>' % (sb, sz, sz, i + 1)
            for i in range(8))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + NS + ">"
                + body + "</w:ftr>")
    body = "".join(para("F%d" % (i + 1)) for i in range(npara))
    if nrows:
        rows = "".join(
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>%s</w:tc>'
            '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr>'
            % (para("R%dA" % (r + 1)), para("R%dB" % (r + 1)))
            for r in range(nrows))
        body += ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
                 '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="3000"/></w:tblGrid>'
                 + rows + "</w:tbl>" + para(""))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + NS + ">"
            + body + "</w:ftr>")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    sect_parts = []
    for ai, (name, npara, nrows) in enumerate(ARMS):
        body.append(para("M%02d" % ai, pbb=ai > 0))
        for k in range(BODY_LINES):
            body.append(para("a%dL%02d" % (ai, k + 1)))
        # each arm is its own section so it can carry its own footer
        sect = ('<w:p><w:pPr><w:sectPr>'
                '<w:footerReference w:type="default" r:id="rIdF%d"/>'
                '<w:pgSz w:w="%d" w:h="%d"/>'
                '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
                'w:header="720" w:footer="%d" w:gutter="0"/>'
                "</w:sectPr></w:pPr></w:p>"
                % (ai, PGW, PGH, MAR_TOP, MAR_LR, MAR_BOT, MAR_LR, FTR_DIST))
        if ai < len(ARMS) - 1:
            body.append(sect)
        else:
            sect_parts.append(
                '<w:sectPr><w:footerReference w:type="default" r:id="rIdF%d"/>'
                '<w:pgSz w:w="%d" w:h="%d"/>'
                '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
                'w:header="720" w:footer="%d" w:gutter="0"/></w:sectPr>'
                % (ai, PGW, PGH, MAR_TOP, MAR_LR, MAR_BOT, MAR_LR, FTR_DIST))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) + "".join(sect_parts) + "</w:body></w:document>")
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
              '<w:sz w:val="20"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
              ' w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style>'
              '<w:style w:type="paragraph" w:styleId="FtrSB"><w:name w:val="FtrSB"/>'
              '<w:basedOn w:val="Normal"/><w:pPr>'
              '<w:spacing w:before="120" w:after="0"/></w:pPr></w:style>'
              '<w:style w:type="paragraph" w:styleId="FtrSA"><w:name w:val="FtrSA"/>'
              '<w:basedOn w:val="Normal"/><w:pPr>'
              '<w:spacing w:before="0" w:after="120"/></w:pPr></w:style>'
              '<w:style w:type="paragraph" w:styleId="FtrSBSA"><w:name w:val="FtrSBSA"/>'
              '<w:basedOn w:val="Normal"/><w:pPr>'
              '<w:spacing w:before="120" w:after="120"/></w:pPr></w:style>'
              "</w:styles>")
    ct_extra = "".join(
        '<Override PartName="/word/footer%d.xml" ContentType='
        '"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' % ai
        for ai in range(len(ARMS)))
    ct_extra += ('<Override PartName="/docProps/core.xml" ContentType='
                 '"application/vnd.openxmlformats-package.core-properties+xml"/>'
                 '<Override PartName="/docProps/app.xml" ContentType='
                 '"application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                 '<Override PartName="/docProps/custom.xml" ContentType='
                 '"application/vnd.openxmlformats-officedocument.custom-properties+xml"/>')
    ct = CT.replace("</Types>", ct_extra + "</Types>")
    rel_extra = "".join(
        '<Relationship Id="rIdF%d" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/footer" Target="footer%d.xml"/>' % (ai, ai)
        for ai in range(len(ARMS)))
    drels = DRELS.replace("</Relationships>", rel_extra + "</Relationships>")
    rels = RELS.replace("</Relationships>",
                        '<Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/'
                        'package/2006/relationships/metadata/core-properties" '
                        'Target="docProps/core.xml"/>'
                        '<Relationship Id="rIdApp" Type="http://schemas.openxmlformats.org/'
                        'officeDocument/2006/relationships/extended-properties" '
                        'Target="docProps/app.xml"/>'
                        '<Relationship Id="rIdCustom" Type="http://schemas.openxmlformats.org/'
                        'officeDocument/2006/relationships/custom-properties" '
                        'Target="docProps/custom.xml"/>'
                        "</Relationships>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/styles.xml", styles)
        for ai, (_n, npara, nrows) in enumerate(ARMS):
            z.writestr("word/footer%d.xml" % ai, footer_xml(npara, nrows))
        # ★the specimen's footer uses DOCPROPERTY fields; without its docProps
        # parts Word renders "Error! Property name is incorrect" strings that
        # WRAP and make the arm's footer taller than the real one — the first
        # run of f9_real measured a 96pt stack that way, 43pt of which was the
        # error text. Copy the parts so the fields resolve.
        zs = zipfile.ZipFile(SPECIMEN)
        for part in [n for n in zs.namelist() if n.startswith("docProps")]:
            z.writestr(part, zs.read(part))
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms | page", PGH / 20.0,
          "pt | bottom margin", MAR_BOT / 20.0, "| footer dist", FTR_DIST / 20.0)


def report(per, who):
    print("== %s ==  (page %.1fpt, declared body bottom %.1fpt)"
          % (who, PGH / 20.0, (PGH - MAR_BOT) / 20.0))
    print("%-11s %-6s %-6s %10s %10s %10s"
          % ("arm", "paras", "rows", "last body", "footer top", "body bottom"))
    for ai, (name, npara, nrows) in enumerate(ARMS):
        got = per.get(ai)
        if not got:
            print("%-11s %-6d %-6d MISSING" % (name, npara, nrows))
            continue
        last_body, ftr_top = got
        print("%-11s %-6d %-6d %10s %10s %10s"
              % (name, npara, nrows,
                 "%.1f" % last_body if last_body else "?",
                 "%.1f" % ftr_top if ftr_top else "?",
                 "%.1f" % (last_body + 10.0) if last_body else "?"))


def pdf():
    import fitz
    import win32com.client as w
    out = docx().replace(".docx", ".pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx(), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(out, 17)
    finally:
        d.Close(False)
        app.Quit()
    doc = fitz.open(out)
    per = {}
    for ai in range(len(ARMS)):
        body_y, ftr_y = None, None
        for pi in range(doc.page_count):
            rows = []
            for bl in doc[pi].get_text("dict")["blocks"]:
                if bl["type"] != 0:
                    continue
                for ln in bl["lines"]:
                    t = "".join(s["text"] for s in ln["spans"]).strip()
                    if t:
                        rows.append((round(ln["bbox"][1], 1), t))
            if not any(t.startswith("a%dL" % ai) for _y, t in rows):
                continue
            body = [y for y, t in rows if t.startswith("a%dL" % ai)]
            ftr = [y for y, t in rows if t.startswith(("F", "R"))]
            if body:
                body_y = max(body_y or 0, max(body))
            if ftr:
                ftr_y = min(ftr_y or 1e9, min(ftr))
        per[ai] = (body_y, ftr_y)
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "ftrheight_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "fh"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    per = {}
    for ai in range(len(ARMS)):
        body_y, ftr_y = None, None
        for pg in pages:
            rows = [(round(e["y"], 1), (e.get("text") or "").strip())
                    for e in pg["elements"] if e.get("type") == "text"]
            if not any(t.startswith("a%dL" % ai) for _y, t in rows):
                continue
            body = [y for y, t in rows if t.startswith("a%dL" % ai)]
            ftr = [y for y, t in rows if t.startswith(("F", "R"))]
            if body:
                body_y = max(body_y or 0, max(body))
            if ftr:
                ftr_y = min(ftr_y or 1e9, min(ftr))
        per[ai] = (body_y, ftr_y)
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
