# -*- coding: utf-8 -*-
"""Two text-anchored floating tables in a row: does the bottom rule of float A count in the
position of float B, whose anchor paragraph follows the paragraph A pushed below itself?

ja/forms__02157d72 p1: float A (tblpY 35, tcBorders sz 4) ends at the rule 302.81; the
paragraph after it («※電話番号», A's anchor) is pushed below A, but float B (anchored to
the next paragraph, tblpYSpec=bottom vertAnchor=text) starts at 316.37 = 302.81 + 13.56,
one ＭＳ 明朝 10.5 line and NO rule width, and «※電話番号» finally lands below B at B's
rule + 0.48. `_pb_floattbl_bottom_gen.py` (one float, its anchor pushed below) measured
rule + width. So the pushed paragraph's own position takes the width, the anchor chain
of the next float may not.

Sheet: «前», float A (two exact 340tw rows, every cell tcBorders sz 4, tblpY = A_Y twips,
tblpXSpec center, tblOverlap never, full width), «中» (A's anchor), [float B (two exact
rows, same borders, vertAnchor text, B_POS)], «後». Arms:
    a_only       no B                                  («中» = rule_A + width?)
    a_b0         B tblpY=0                             (B top = ? ; «中» below B)
    a_bbottom    B tblpYSpec=bottom                    (the document's form)
    a80_b0       A tblpY=80, B tblpY=0
    bottom_a12 / bottom_b12 / b0_a12   A or B with sz 12 rules (what the bottom shift is made of)
    bonly_bottom / bonly_0             no A at all
    bonly_top / center / inside / outside   the other tblpYSpec values (text anchor)
Read: baselines of 前 / 中 / 後 / 一 (A row 1) / 三 (B row 1), horizontal rules.

    python tools/metrics/_pb_float_chain_gen.py            (PB_OXI=<exe>[;<exe>] adds Oxi; PB_ARMS=a,b limits; PB_SZ / PB_ROWS / PB_GRID=none vary the sheet)
"""
import json, os, subprocess, sys, tempfile, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_float_chain")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:compat><w:useFELayout/>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % W
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="ＭＳ 明朝"/><w:kern w:val="2"/><w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>') % W
PB_SZ = os.environ.get("PB_SZ")
PB_ROWS = int(os.environ.get("PB_ROWS", "2"))
SUFFIX = ("_sz" + PB_SZ if PB_SZ else "") + ("_r%d" % PB_ROWS if PB_ROWS != 2 else "") + ("_nogrid" if os.environ.get("PB_GRID") == "none" else "")


def tcb(sz):
    return '<w:tcBorders>' + "".join('<w:%s w:val="single" w:sz="%d" w:space="0" w:color="auto"/>' % (e, sz) for e in ("top", "left", "bottom", "right")) + '</w:tcBorders>'


def table(pos_attr, texts, sz=4):
    t = ('<w:tbl><w:tblPr><w:tblpPr w:leftFromText="142" w:rightFromText="142" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="center" %s/>'
         '<w:tblOverlap w:val="never"/><w:tblW w:w="10000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
         '<w:tblCellMar><w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="10000"/></w:tblGrid>' % pos_attr)
    for x in texts:
        t += ('<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="340"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="10000" w:type="dxa"/>%s</w:tcPr>'
              '<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr><w:r><w:t>%s</w:t></w:r></w:p></w:tc></w:tr>' % (tcb(sz), x))
    return t + "</w:tbl>"


def build(name, a_y, b_pos, a_sz=4, b_sz=4):
    # PB_SZ (half-points) sizes every body run and mark; PB_ROWS rows in B; PB_GRID=none drops the docGrid
    rpr = '<w:rPr><w:sz w:val="%s"/></w:rPr>' % PB_SZ if PB_SZ else ""
    sp = '<w:pPr><w:snapToGrid w:val="0"/>%s</w:pPr>' % rpr
    para = lambda s: '<w:p>%s<w:r>%s<w:t>%s</w:t></w:r></w:p>' % (sp, rpr, s)
    body = para("前") + (table('w:tblpY="%d"' % a_y, ["一", "二"], a_sz) if a_y is not None else "") + para("中")
    if b_pos is not None:
        body += table(b_pos, (["三", "四", "五", "六", "七", "八"])[:PB_ROWS], b_sz)
    body += para("後")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="567" w:right="567" w:bottom="284" w:left="1134" w:header="851" w:footer="992"/>'
           '%s</w:sectPr></w:body></w:document>') % (W, body, "" if os.environ.get("PB_GRID") == "none" else '<w:docGrid w:type="linesAndChars" w:linePitch="290" w:charSpace="-1209"/>')
    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, name + ".docx")
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES); z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    return p


KEYS = "前一中三後"


def read_word(pdf):
    import pymupdf
    pg = pymupdf.open(pdf)[0]
    ys = {}
    for b in pg.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"] in KEYS:
                        ys.setdefault(c["c"], c["origin"][1])
    rules = sorted({round(d["rect"].y0, 2) for d in pg.get_drawings() if d["rect"].width > 100 and d["rect"].height < 3})
    return ys, rules


def read_oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json"); lay = os.path.join(t, "l.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    subprocess.run([exe, docx, os.path.join(t, "p2"), "--dump-layout=" + lay], capture_output=True)
    ys = {}
    for c in json.load(open(g, encoding="utf-8"))["pages"][0]["glyphs"]:
        if c["char"] in KEYS:
            ys.setdefault(c["char"], c["baseline"])
    rules = sorted({round(e["y"], 2) for e in json.load(open(lay, encoding="utf-8"))["pages"][0]["elements"]
                    if e.get("type") == "border" and e.get("w", 0) > 100})
    return ys, rules


def fmt(ys, rules):
    return " ".join("%s %.2f" % (k, ys[k]) for k in KEYS if k in ys) + "  rules %s" % rules


if __name__ == "__main__":
    import win32com.client
    exe = os.environ.get("PB_OXI")
    arms = [("a_only", 35, None, 4, 4), ("a_b0", 35, 'w:tblpY="0"', 4, 4), ("a_bbottom", 35, 'w:tblpYSpec="bottom"', 4, 4), ("a80_b0", 80, 'w:tblpY="0"', 4, 4),
            # second round: what the yspec=bottom shift is made of (A's rule, B's rule, or neither)
            ("bottom_a12", 35, 'w:tblpYSpec="bottom"', 12, 4), ("bottom_b12", 35, 'w:tblpYSpec="bottom"', 4, 12), ("b0_a12", 35, 'w:tblpY="0"', 12, 4),
            ("bonly_bottom", None, 'w:tblpYSpec="bottom"', 4, 4), ("bonly_0", None, 'w:tblpY="0"', 4, 4),
            ("bonly_top", None, 'w:tblpYSpec="top"', 4, 4), ("bonly_center", None, 'w:tblpYSpec="center"', 4, 4),
            ("bonly_inside", None, 'w:tblpYSpec="inside"', 4, 4), ("bonly_outside", None, 'w:tblpYSpec="outside"', 4, 4)]
    if os.environ.get("PB_ARMS"):
        arms = [a for a in arms if a[0] in os.environ["PB_ARMS"].split(",")]
    for name, a_y, b_pos, a_sz, b_sz in arms:
        p = build(name + SUFFIX, a_y, b_pos, a_sz, b_sz); pdf = p[:-5] + ".pdf"
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
        print("%-10s W " % (name + SUFFIX) + fmt(*read_word(pdf)), flush=True)
        if exe:
            for e in exe.split(";"):
                print("%-10s O " % os.path.basename(e)[12:-4] + fmt(*read_oxi(e, p)), flush=True)
