# -*- coding: utf-8 -*-
"""Does a SPLIT TABLE ROW have to contain the boundary paragraph's space_after?

S1092 (2026-08-07) was derived on policies__0028d1be: a paragraph whose box
fits by 11.7pt is PUSHED because its 14pt after does not fit (an after=0 arm
KEEPS the identical geometry with 3.7pt of room).  But uklocalspending's
single-paragraph template cells KEEP with the same shape, and the structural
difference is that there the boundary paragraph is the LAST one in its cell.

This probe tests that axis directly.  Each arm is a one-cell table row that
spans two pages; the cell holds N filler lines, then the TARGET paragraph
(after = A), and in the MID arms one more paragraph after it.  Sweeping the
bottom margin moves the content bottom across the target's box:

    LAST arm  : flip expected when box_bottom          > cbot
    MID  arm  : flip expected when box_bottom + A      > cbot   (A earlier)

so the two flip points differ by exactly A if the after is required only when
another paragraph follows inside the same cell.

  python _pb_rowsplit_after_gen.py gen | bake | read

MEASURED (24 arms, Word, bottom margin in 2pt steps): LAST and MID flip at the
SAME point, cbot ~= 734 = the target's box bottom 720.3 + its after 14.  So the
after is required in BOTH positions -- the "last paragraph in the cell is
exempt" narrowing is FALSIFIED.  Combined with uklocalspending's render-truth
(the row's bottom border sits at last_line_bottom + q, not + after + q, where
q = tcMar_b + border), the rule is:

    a split-row fragment closes at  last_line_bottom + max(after, q)

i.e. the trailing after and the cell's bottom frame OVERLAP rather than stack.
"""
import os, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.abspath(os.path.join(HERE, "..", "..", "pipeline_data", "_pb_rowsplit_after"))
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
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>'
          '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '</w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '</w:styles>')

NFILL = 45          # filler lines inside the cell (TNR 12pt = 13.8pt each)
                    # target box = 706.8 .. 720.6; +after 14 -> 734.6
AFTER = 280         # twips = 14.0pt, the value policies__0028d1be uses
BOTTOMS = list(range(1100, 1541, 40))   # 55 .. 77pt in 2pt steps


def p(text, after=0, mark=""):
    return ('<w:p><w:pPr><w:spacing w:after="%d" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:t xml:space="preserve">%s%s</w:t></w:r></w:p>' % (after, mark, text))


NFILL_B = 50        # the SECOND cell is long, so the row always splits far below


def row(mid):
    """A TWO-cell row: cell A ends at TARGET (+ TAIL in the MID arms) while cell B
    keeps running, so the row must split regardless of where TARGET lands.  This
    separates the split-fit rule from the trivial `does the whole row fit`."""
    a = "".join(p("a%02d" % i) for i in range(NFILL))
    a += p("TARGET", after=AFTER)
    if mid:
        a += p("TAIL")
    b = "".join(p("b%02d" % i) for i in range(NFILL_B))
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>'
            '<w:tr>'
            '<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>%s</w:tc>'
            '<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>%s</w:tc>'
            '</w:tr></w:tbl>' % (a, b))


def sect(bottom, last):
    s = ('<w:sectPr>%s<w:pgSz w:w="12240" w:h="15840"/>'
         '<w:pgMar w:top="1440" w:right="1440" w:bottom="%d" w:left="1440" '
         'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
         % ("" if last else '<w:type w:val="nextPage"/>', bottom))
    return s if last else '<w:p><w:pPr>%s</w:pPr></w:p>' % s


def build(path, arms):
    parts = []
    for i, (name, bottom, mid) in enumerate(arms):
        parts.append(p("=== %s ===" % name))
        parts.append(row(mid))
        parts.append(sect(bottom, i == len(arms) - 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           '><w:body>' + "".join(parts) + '</w:body></w:document>')
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)


def arms():
    out = []
    for b in BOTTOMS:
        out.append(("LAST_%d" % b, b, False))
    for b in BOTTOMS:
        out.append(("MID_%d" % b, b, True))
    return out


def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    os.makedirs(OUT, exist_ok=True)
    A = arms()
    docx = os.path.join(OUT, "rowsplit_after.docx")
    if cmd == "gen":
        build(docx, A)
        print("built %s  (%d arms)" % (docx, len(A)))
    elif cmd == "read":
        import win32com.client as w
        app = w.DispatchEx("Word.Application"); app.Visible = False
        try:
            doc = app.Documents.Open(docx, ReadOnly=True); doc.Repaginate()
            marks = []
            for i in range(1, doc.Paragraphs.Count + 1):
                r = doc.Paragraphs(i).Range
                t = r.Text.replace("\r", "").replace("\x07", "")
                if t.startswith("=== ") or t == "TARGET":
                    c = doc.Range(r.Start, r.Start)
                    marks.append((t, c.Information(3), c.Information(6)))
            doc.Close(False)
        finally:
            app.Quit()
        cur, res = None, {}
        for t, pg, y in marks:
            if t.startswith("==="):
                cur = t.strip("= ").strip()
            elif cur:
                res.setdefault(cur, (pg, y))
        prev = {}
        for name, bottom, mid in A:
            pg, y = res.get(name, (None, None))
            cbot = 792.0 - bottom / 20.0
            kind = "MID " if mid else "LAST"
            # pushed == the target restarts at the page top (72 + cell inset)
            pushed = (y is not None and y < 200.0)
            flip = ""
            if kind in prev and prev[kind] != pushed:
                flip = "  <<< FLIP"
            prev[kind] = pushed
            print("%-5s bottom=%-6.2f cbot=%-8.2f y=%-8s %s%s"
                  % (kind, bottom / 20.0, cbot, ("%.2f" % y) if y else "-",
                     "PUSH" if pushed else "KEEP", flip))
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
