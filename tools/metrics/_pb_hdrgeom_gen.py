"""§9 probe for REPORT_reference__00215c: pin the body-start model under a
border-only header (Stage B of the per-type-inheritance fix).

Target header2.xml = [bordered para before=1000 line=240 border sz=12 space=1]
+ [trailing empty Header-style para]. Word's continuation-page body starts
~4pt BELOW the header border bottom (border bottom 101.18, body needs ~105).
The ~3.8pt component is one of: border ink extent / header flow cursor (+1 line
for the trailing empty) / paragraph-mark / docGrid snap.

This probe puts the header DIRECTLY in each section (no inheritance needed to
measure geometry) and measures, per config, Word's:
  - header border bottom y
  - body first line y (continuation page)
across: trailing-empty {no,yes} x docGrid {none,360} x before {500,1000} x
border {sz 4, 12}. The differences pin the model.

Usage: python _pb_hdrgeom_gen.py gen | measure | read
"""
import os, sys, json, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_hdrgeom")

# config axes: (label, trailing_empty, grid_lp, before_tw, border_sz)
CONFIGS = [
    ("base",      False, 360, 1000, 12),   # target-like WITHOUT trailing empty
    ("trail",     True,  360, 1000, 12),   # target-like WITH trailing empty
    ("trail_nog", True,  0,   1000, 12),   # trailing, no docGrid
    ("base_nog",  False, 0,   1000, 12),   # no trailing, no docGrid
    ("trail_b500",True,  360, 500,  12),   # smaller before
    ("trail_sz4", True,  360, 1000, 4),    # thinner border
]

HDR_RID = "rId50"


def header_xml(trailing, before_tw, border_sz):
    bordered = (f'<w:p><w:pPr>'
                f'<w:pBdr><w:bottom w:val="single" w:sz="{border_sz}" w:space="1" w:color="auto"/></w:pBdr>'
                f'<w:spacing w:before="{before_tw}" w:line="240" w:lineRule="auto"/>'
                f'</w:pPr></w:p>')
    trail = '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p>' if trailing else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + bordered + trail + '</w:hdr>')


def body(grid_lp):
    # 55 body lines to force multiple pages; the section's own default header.
    paras = ''.join(f'<w:p><w:r><w:t>Body line {i} of the section content here now.</w:t></w:r></w:p>'
                    for i in range(55))
    grid = f'<w:docGrid w:linePitch="{grid_lp}"/>' if grid_lp else ''
    sect = (f'<w:sectPr>'
            f'<w:headerReference w:type="default" r:id="{HDR_RID}"/>'
            f'<w:pgSz w:w="11907" w:h="16839"/>'
            f'<w:pgMar w:top="1843" w:right="1440" w:bottom="1440" w:left="1440" '
            f'w:header="720" w:footer="720" w:gutter="0"/>{grid}</w:sectPr>')
    return paras + sect


def write_docx(path, trailing, grid_lp, before_tw, border_sz):
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           f'<w:body>{body(grid_lp)}</w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/header50.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             f'<Relationship Id="{HDR_RID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header50.xml"/>'
             '</Relationships>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
              '<w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/></w:style>'
              '</w:styles>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/header50.xml", header_xml(trailing, before_tw, border_sz))


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for (lab, te, lp, bf, sz) in CONFIGS:
        write_docx(os.path.join(OUTDIR, f"h_{lab}.docx"), te, lp, bf, sz)
    print(f"wrote {len(CONFIGS)} docs")


def measure():
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False; word.DisplayAlerts = 0
    try:
        for (lab, *_r) in CONFIGS:
            src = os.path.abspath(os.path.join(OUTDIR, f"h_{lab}.docx"))
            d = word.Documents.Open(src, ReadOnly=True)
            d.ExportAsFixedFormat(src[:-5] + ".pdf", 17)
            d.Close(False)
        print("exported")
    finally:
        word.Quit()


def read():
    import fitz
    print(f"{'cfg':11} {'border_y1':>9} {'body1_y':>8} {'gap':>6} {'note'}")
    rows = []
    for (lab, te, lp, bf, sz) in CONFIGS:
        pdf = os.path.join(OUTDIR, f"h_{lab}.pdf")
        if not os.path.exists(pdf):
            continue
        doc = fitz.open(pdf)
        pg = doc[len(doc) - 1] if len(doc) > 1 else doc[0]  # a continuation page
        # header border = a thin horizontal drawing near y~100; body first glyph
        border_y1 = None
        for dr in pg.get_drawings():
            for it in dr["items"]:
                if it[0] == "l":  # line
                    (x0, y0), (x1, y1) = it[1], it[2]
                    if abs(y0 - y1) < 0.5 and 90 < y0 < 115 and abs(x1 - x0) > 100:
                        border_y1 = max(border_y1 or 0, y0)
                elif it[0] == "re":  # rect (thin border)
                    r = it[1]
                    if r.height < 3 and 90 < r.y0 < 115 and r.width > 100:
                        border_y1 = max(border_y1 or 0, r.y1)
        body1_y = None
        for blk in pg.get_text("dict")["blocks"]:
            for ln in blk.get("lines", []):
                t = "".join(s["text"] for s in ln["spans"])
                if "Body line" in t:
                    y = min(s["bbox"][1] for s in ln["spans"])
                    if body1_y is None or y < body1_y:
                        body1_y = y
        doc.close()
        gap = (body1_y - border_y1) if (body1_y and border_y1) else None
        print(f"{lab:11} {border_y1 or 0:>9.2f} {body1_y or 0:>8.2f} "
              f"{gap if gap is not None else 0:>6.2f} te={te} lp={lp} bf={bf} sz={sz}")
        rows.append({"cfg": lab, "border_y1": border_y1, "body1_y": body1_y,
                     "gap": gap, "trailing": te, "grid": lp, "before": bf, "sz": sz})
    json.dump(rows, open(os.path.join(OUTDIR, "_result.json"), "w"), indent=1)
    print("\nModel discriminators:")
    print("  border ink extent -> body1 = border_y1 + <font descent>, INVARIANT to trailing/grid")
    print("  header flow cursor -> trail adds ~1 Header line vs base")
    print("  docGrid snap       -> gap follows linePitch remainder (base vs base_nog differ)")


if __name__ == "__main__":
    {"gen": gen, "measure": measure, "read": read}[sys.argv[1]]()
