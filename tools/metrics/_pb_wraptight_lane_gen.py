"""§4.4 probe for REPORT_reports__0013bcb8 ROOT A: pin Word's wrapTight
narrow-lane cutoff.

reports__0013bcb8 p1: a gray wrapTight title textbox (397.35x70.8pt) leaves a
~46.7pt right lane. Word sends the following «Research Article» paragraph BELOW
the band (its 2-word natural width ~79.65pt doesn't fit the lane); Oxi puts it
in the lane and splits Research/Article. The escape is only
`max(left_room,right_room) < 30pt`. H2 (longest word fits -> beside) is REFUTED
(Research 43.8pt fits 46.7 but Word sends below).

This probe fixes a wrapTight textbox at the column top-left and sweeps its width
(= the right lane) x the candidate paragraph's natural width. If the flip
lane tracks the candidate's natural 1-line width -> H3 (paragraph natural fit).
If the flip lane is fixed regardless of candidate -> H1 (fixed min lane). If
both -> H4.

Read from Word PDF: the candidate's first span x/y. beside = x in the lane;
below = x at margin AND y >= textbox bottom.

Usage: python _pb_wraptight_lane_gen.py gen | measure | read
"""
import os, sys, json, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_wraptight_lane")

FONT = "Book Antiqua"
EMU = 12700  # per pt
MARGIN = 72.0
PAGE_W = 595.3  # A4 pt
CONTENT_W = PAGE_W - 2 * MARGIN  # 451.3
TB_X = 0.65  # textbox offset from column left (posOffsetH ~8255 EMU)

# candidates: (label, text, approx natural width note)
CANDS = [
    ("short", "In it"),          # 2 tiny words, natural ~20pt
    ("target", "Research Article"),  # the target case, natural ~79.65pt
    ("long", "Research Articles Now"),  # 3 words, natural ~110pt
]
# right-lane widths to sweep (pt) — fine around the 36..48 flip
LANES = [24, 36, 40, 42, 44, 45, 46, 47, 48, 50, 60, 84]


def wps_textbox(cx_emu):
    # A wrapTight PICTURE (the keep-out band geometry is what matters, not the
    # object type — Word rejects a bare wps:txbx without mc:AlternateContent).
    return (
        '<w:r><w:drawing>'
        '<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>8255</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>78740</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx_emu}" cy="899160"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0">'
        '<wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/>'
        '<wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>'
        '<wp:docPr id="1" name="img"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx_emu}" cy="899160"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:anchor>'
        '</w:drawing></w:r>')


# minimal 1x1 gray PNG
_PNG = __import__("base64").b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNsaGj4DwAFhAKA"
    "3wZ+DwAAAABJRU5ErkJggg==")


def build(cand_text, lane_pt):
    # right lane = CONTENT_W - TB_X - cx_pt  ->  cx_pt = CONTENT_W - TB_X - lane
    cx_pt = CONTENT_W - TB_X - lane_pt
    cx_emu = int(round(cx_pt * EMU))
    rpr = f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}"/><w:sz w:val="21"/>'  # 10.5pt
    # anchor paragraph carries the textbox drawing + the candidate text
    anchor = (f'<w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr>'
              f'{wps_textbox(cx_emu)}'
              f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{cand_text}</w:t></w:r></w:p>')
    # a few filler body paras below so the page has content
    filler = ''.join(f'<w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr>'
                     f'<w:r><w:rPr>{rpr}</w:rPr><w:t>Body line {i} lorem ipsum dolor sit amet consectetur.</w:t></w:r></w:p>'
                     for i in range(6))
    mar = (f'<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
           'w:header="720" w:footer="720" w:gutter="0"/>')
    sect = f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>{mar}</w:sectPr>'
    return anchor + filler + sect


def write_docx(path, cand_text, lane_pt):
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{build(cand_text, lane_pt)}</w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Default Extension="png" ContentType="image/png"/>'
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
             '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>'
             '</Relationships>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              f'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>'
              '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '</w:styles>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/media/image1.png", _PNG)


def nm(clabel, lane):
    return f"wt_{clabel}_L{lane}.docx"


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    n = 0
    for clabel, ctext in CANDS:
        for lane in LANES:
            write_docx(os.path.join(OUTDIR, nm(clabel, lane)), ctext, lane)
            n += 1
    print(f"wrote {n} docs to {OUTDIR}")


def measure():
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False; word.DisplayAlerts = 0
    try:
        for clabel, _ in CANDS:
            for lane in LANES:
                src = os.path.abspath(os.path.join(OUTDIR, nm(clabel, lane)))
                d = word.Documents.Open(src, ReadOnly=True)
                d.ExportAsFixedFormat(src[:-5] + ".pdf", 17)
                d.Close(False)
        print("exported PDFs")
    finally:
        word.Quit()


def read():
    import fitz
    print(f"{'cand':8} {'lane':>5} {'cand_x':>7} {'cand_y':>7} {'verdict':>8} {'lines':>5}")
    res = []
    for clabel, ctext in CANDS:
        first_word = ctext.split()[0]
        for lane in LANES:
            pdf = os.path.join(OUTDIR, nm(clabel, lane)[:-5] + ".pdf")
            if not os.path.exists(pdf):
                continue
            doc = fitz.open(pdf)
            pg = doc[0]
            # find the candidate's first span (starts with first_word), and the
            # textbox bottom (the gray rect / "Title Box")
            cand_spans = []
            tb_bottom = 0.0
            for blk in pg.get_text("dict")["blocks"]:
                for ln in blk.get("lines", []):
                    t = "".join(s["text"] for s in ln["spans"])
                    if "Title Box" in t:
                        tb_bottom = max(tb_bottom, max(s["bbox"][3] for s in ln["spans"]))
                    if first_word in t and "Body line" not in t:
                        x0 = min(s["bbox"][0] for s in ln["spans"])
                        y0 = min(s["bbox"][1] for s in ln["spans"])
                        cand_spans.append((round(x0, 2), round(y0, 2), t[:24]))
            doc.close()
            if not cand_spans:
                continue
            cx, cy, _ = cand_spans[0]
            # beside: x well right of margin (in the lane); below: x ~ margin(72) and y >= tb_bottom
            beside = cx > MARGIN + 40
            verdict = "BESIDE" if beside else "BELOW"
            nlines = len(cand_spans)
            print(f"{clabel:8} {lane:>5} {cx:>7.2f} {cy:>7.2f} {verdict:>8} {nlines:>5}")
            res.append({"cand": clabel, "lane": lane, "x": cx, "y": cy,
                        "verdict": verdict, "lines": nlines})
    json.dump(res, open(os.path.join(OUTDIR, "_result.json"), "w"), indent=1)


if __name__ == "__main__":
    {"gen": gen, "measure": measure, "read": read}[sys.argv[1]]()
