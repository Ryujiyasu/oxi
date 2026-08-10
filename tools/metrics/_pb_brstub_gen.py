# -*- coding: utf-8 -*-
"""How much page room does a `<w:br w:type="page"/>`-ONLY paragraph demand?

S1055 makes such a stub "a normal line" and denies it the S736 empty-paragraph
tolerance, derived from legal__001a2c7f (its stub overflowed by 1.805pt and
Word pushed it to a blank page).  policies__000f7115 under S1091 shows the
opposite: its stub overflows Oxi's content bottom by 0.443pt and Word KEEPS it
on the page (Word p5 ends with the body line above it, p6 starts with the next
heading at the very top -- so the stub occupies neither page's ink).

Two readings fit those two documents:
  H  the stub's line is the RAW natural height (the pPr line multiplier does
     NOT apply to a break-only line)
  T  the stub takes the multiplied box but Word tolerates a small overflow

They separate cleanly: at the same font size, an arm whose stub paragraph
declares line=240 (1.0x) and one that inherits docDefaults line=259 (1.0792x)
flip at the SAME cursor under H and 1.27pt apart under T.

Each arm is its own section: N filler lines, an exact-height spacer that walks
the cursor in 0.5pt steps, the stub, then a MARKER paragraph.  KEEP puts the
marker one page after the last filler, PUSH puts it two pages after.

  python _pb_brstub_gen.py gen | read [fine]
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_brstub")
DOCX = os.path.join(OUT, "brstub.docx")
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
# Same shape as policies__000f7115: Arial body, docDefaults line=259 auto.
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/>'
          '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="259" w:lineRule="auto"/></w:pPr>'
          '</w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/></w:style></w:styles>')

NLINES = 45

# (key, stub half-point size, stub pPr spacing override or "")
CONFIGS = [
    ("A", 28, ""),                                            # 14pt, docDefaults line=259
    ("B", 28, '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'),   # 14pt, 1.0x
    ("C", 40, ""),                                            # 20pt, docDefaults line=259
]
COARSE = list(range(700, 1160, 40))
# Coarse put A/B's flip in (47, 49] and C's in (39, 41] -- 0.5pt steps over
# both windows separate raw hhea (A==B, C 6.90 earlier) from the multiplied
# box (A 1.27 before B, C 7.44 earlier).
FINE = list(range(760, 1010, 10))


def txt(s):
    return ('<w:p><w:pPr><w:spacing w:after="0" w:line="259" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % s)


def spacer(x):
    return ('<w:p><w:pPr><w:spacing w:after="0" w:line="%d" w:lineRule="exact"/></w:pPr></w:p>'
            % x)


def stub(sz, sp):
    rpr = ('<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
           '<w:b/><w:sz w:val="%d"/></w:rPr>' % sz)
    return ('<w:p><w:pPr>%s%s</w:pPr><w:r>%s<w:br w:type="page"/></w:r></w:p>'
            % (sp, rpr, rpr))


def sect(last):
    s = ('<w:sectPr>%s<w:pgSz w:w="11906" w:h="16838"/>'
         '<w:pgMar w:top="720" w:right="720" w:bottom="1440" w:left="720" '
         'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
         % ("" if last else '<w:type w:val="nextPage"/>'))
    return s if last else '<w:p><w:pPr>%s</w:pPr></w:p>' % s


def arms(xs):
    return [(k, sz, sp, x) for (k, sz, sp) in CONFIGS for x in xs]


def gen(xs):
    os.makedirs(OUT, exist_ok=True)
    a = arms(xs)
    body = []
    for i, (k, sz, sp, x) in enumerate(a):
        tag = "%s%03d" % (k, x)
        for j in range(NLINES):
            body.append(txt("L%02d_%s" % (j, tag)))
        body.append(spacer(x))
        body.append(stub(sz, sp))
        body.append(txt("TGT_%s" % tag))
        body.append(sect(i == len(a) - 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           '><w:body>' + "".join(body) + "</w:body></w:document>")
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(a), "arms")


def read(xs):
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    rows = {}
    try:
        d.Repaginate()
        for i in range(1, d.Paragraphs.Count + 1):
            r = d.Paragraphs(i).Range
            t = r.Text.strip()
            if not (t.startswith("L%02d_" % (NLINES - 1)) or t.startswith("TGT_")):
                continue
            c = d.Range(r.Start, r.Start)
            tag = t.split("_", 1)[1]
            rows.setdefault(tag, {})["L" if t.startswith("L") else "T"] = (
                c.Information(3), c.Information(6))
    finally:
        d.Close(False)
        app.Quit()
    print("%-6s %8s %9s %8s  %s" % ("arm", "spacer", "cursor", "dpage", "verdict"))
    for k, sz, _sp, x in arms(xs):
        tag = "%s%03d" % (k, x)
        r = rows.get(tag)
        if not r or "L" not in r or "T" not in r:
            print("%-6s MISSING" % tag)
            continue
        # cursor before the stub = last filler's box top + its line + the spacer
        cur = r["L"][1] + x / 20.0
        print("%-6s %8.2f %9.2f %8d  %s"
              % (tag, x / 20.0, cur, r["T"][0] - r["L"][0],
                 "KEEP" if r["T"][0] - r["L"][0] == 1 else "PUSH"))


if __name__ == "__main__":
    cmd = sys.argv[1]
    xs = FINE if len(sys.argv) > 2 and sys.argv[2] == "fine" else COARSE
    {"gen": gen, "read": read}[cmd](xs)
