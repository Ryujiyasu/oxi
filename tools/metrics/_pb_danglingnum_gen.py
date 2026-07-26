"""§10 probe for REPORT_reference__002040a2: does a dangling-numId list's body
indent track defaultTabStop, or is it a fixed 36pt gallery recovery?

reference__002040a2 references numId=1 with an EMPTY numbering.xml (num=0 only) +
a missing ListParagraph style. Word recovers with a decimal marker "1."..."7." at
body-left = margin + 36pt. 36pt == the doc's defaultTabStop (720tw). The target
can't distinguish "= defaultTabStop" from "= fixed 36" (its tabstop IS 36). This
probe sweeps defaultTabStop {480, 720, 960 tw} to pin the body-left model.

Read from Word PDF: marker text (decimal?) + first-text x + continuation x +
line count. If body-left tracks 24/36/48pt -> model A (defaultTabStop). If always
36 -> model B (fixed gallery). If no marker/margin -> model E.

Usage: python _pb_danglingnum_gen.py gen | measure | read
"""
import os, sys, json, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_danglingnum")

ITEM = ("LONG ITEM {n:02d} with enough text to wrap onto a second line when the "
        "list body is indented by the recovered tab stop amount from the margin")

# variants: (label, default_tab_stop_tw, num_id, pstyle)
CASES = [
    ("D1_480", 480, "1", "ListParagraph"),   # dangling, missing style
    ("D2_720", 720, "1", "ListParagraph"),   # = target
    ("D3_960", 960, "1", "ListParagraph"),
    ("D4_num7", 720, "7", "ListParagraph"),  # numId!=1
    ("D5_normal", 720, "1", "Normal"),       # pStyle=Normal (not List)
]


def body(num_id, pstyle):
    ps = []
    for i in range(1, 8):
        ps.append(
            f'<w:p><w:pPr><w:pStyle w:val="{pstyle}"/>'
            f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>'
            f'<w:r><w:t>{ITEM.format(n=i)}</w:t></w:r></w:p>')
    mar = ('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
           'w:header="720" w:footer="720" w:gutter="0"/>')
    sect = f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>{mar}</w:sectPr>'
    return "".join(ps) + sect


def write_docx(path, num_id, pstyle, dts_tw):
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{body(num_id, pstyle)}</w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
             '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             '</Relationships>')
    # styles.xml: Normal (Calibri 11), NO ListParagraph (missing style) for List variants
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/>'
              '</w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '</w:styles>')
    # numbering.xml: EMPTY (dangling numId=1/7)
    numbering = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'<w:defaultTabStop w:val="{dts_tw}"/></w:settings>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/numbering.xml", numbering)
        z.writestr("word/settings.xml", settings)


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for label, dts, num_id, pstyle in CASES:
        write_docx(os.path.join(OUTDIR, f"{label}.docx"), num_id, pstyle, dts)
    print(f"wrote {len(CASES)} docs to {OUTDIR}")


def measure():
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False; word.DisplayAlerts = 0
    try:
        for label, *_ in CASES:
            src = os.path.abspath(os.path.join(OUTDIR, f"{label}.docx"))
            d = word.Documents.Open(src, ReadOnly=True)
            d.ExportAsFixedFormat(src[:-5] + ".pdf", 17)
            d.Close(False)
        print(f"exported {len(CASES)} PDFs")
    finally:
        word.Quit()


def read():
    import fitz
    print(f"{'case':10} {'dts_pt':>6} {'marker':>10} {'body_x':>7} {'cont_x':>7} {'lines(first2)':>14}")
    for label, dts, num_id, pstyle in CASES:
        pdf = os.path.join(OUTDIR, f"{label}.pdf")
        if not os.path.exists(pdf):
            continue
        doc = fitz.open(pdf)
        pg = doc[0]
        # collect the first list item's line(s): text starting with "LONG ITEM 01"
        item1 = []
        marker = None
        for blk in pg.get_text("dict")["blocks"]:
            for ln in blk.get("lines", []):
                spans = ln["spans"]
                t = "".join(s["text"] for s in spans)
                if "LONG ITEM 01" in t or (item1 and "LONG ITEM" not in t and item1[-1][0] and len(item1) < 2):
                    x0 = min(s["bbox"][0] for s in spans)
                    item1.append((t[:20], round(x0, 2)))
                # marker: a short span (decimal "1." or bullet) left of the body
                if "LONG ITEM 01" in t:
                    # the first span's text may be the marker if separate
                    marker = spans[0]["text"][:6]
        doc.close()
        body_x = item1[0][1] if item1 else None
        cont_x = item1[1][1] if len(item1) > 1 else None
        dts_pt = dts / 20.0
        print(f"{label:10} {dts_pt:>6.1f} {str(marker):>10} "
              f"{body_x if body_x else 0:>7.2f} {cont_x if cont_x else 0:>7.2f} "
              f"{str([i[0][:8] for i in item1[:2]]):>14}")


if __name__ == "__main__":
    {"gen": gen, "measure": measure, "read": read}[sys.argv[1]]()
