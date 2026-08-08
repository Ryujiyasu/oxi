# -*- coding: utf-8 -*-
"""Derive Word's EFFECTIVE right edge X for a NON-justified paragraph.

Model under test:  a line keeps its last word iff
    line_natural_width <= (content_width - ind_left - ind_right) - X

One section per arm (1 page each), ind_right swept in 10tw (0.5pt) steps.
The flip (1 line -> 2 lines) pins X to 0.5pt.

usage:  gen | measure | read
"""
import sys, os, zipfile, shutil, subprocess, json, re
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = os.path.abspath(os.path.join("pipeline_data", "_pb_rightedge2"))
DOCX = os.path.join(OUT, "re2.docx")
PDF = os.path.join(OUT, "re2.pdf")

# Letter, the target's most common section geometry
PGW, PGH, ML, MR = 12240, 15840, 1440, 1460
CONTENT = PGW - ML - MR            # 9340tw = 467.0pt
IND_L = 720                        # 36.0pt
# TNR 12pt advances (hmtx/2048*12)
W_M, W_SP, W_T, W_O = 10.6699, 3.0, 3.334, 6.0
NM = 38                            # 38 'M' + ' to'
BODY_W = NM * W_M + W_SP + W_T + W_O          # natural width of the whole line
HEAD_W = NM * W_M                             # width without the last word ' to'
ARMS = list(range(0, 40))          # ind_right = 0,10,...,390 tw  (0..19.5pt)

def sect(ir, last=False):
    t = "" if last else '<w:type w:val="nextPage"/>'
    return (f'<w:sectPr>{t}<w:pgSz w:w="{PGW}" w:h="{PGH}"/>'
            f'<w:pgMar w:top="1440" w:right="{MR}" w:bottom="1440" w:left="{ML}" '
            f'w:header="720" w:footer="720" w:gutter="0"/>'
            f'<w:cols w:space="0"/></w:sectPr>')

def para(txt, ir, tag):
    return (f'<w:p><w:pPr><w:ind w:left="{IND_L}" w:right="{ir}"/>'
            f'<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{tag} {txt}</w:t></w:r></w:p>')

def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for i, k in enumerate(ARMS):
        ir = k * 10
        tag = f"A{i:02d}"
        # the tag adds width, so shorten the M-run by its measured cost later;
        # we instead put the tag on its OWN paragraph so the measured line is pure.
        body.append(f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                    f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
                    f'<w:sz w:val="24"/></w:rPr><w:t>{tag}</w:t></w:r></w:p>')
        body.append(f'<w:p><w:pPr><w:ind w:left="{IND_L}" w:right="{ir}"/>'
                    f'<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                    f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
                    f'<w:sz w:val="24"/></w:rPr>'
                    f'<w:t xml:space="preserve">{"M"*NM} to</w:t></w:r></w:p>')
        body.append(f'<w:p><w:pPr>{sect(ir, last=(i == len(ARMS)-1))}</w:pPr></w:p>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>' + "".join(body) + '</w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
    print("gen", DOCX, "arms", len(ARMS), "BODY_W", round(BODY_W, 3), "HEAD_W", round(HEAD_W, 3))

def measure():
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application"); app.Visible = False
    try:
        d = app.Documents.Open(DOCX, ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=PDF, ExportFormat=17)
        d.Close(False)
    finally:
        app.Quit()
    print("measure", PDF)

def read():
    import fitz
    d = fitz.open(PDF)
    print(f"content={CONTENT/20:.1f}pt ind_left={IND_L/20:.1f}pt BODY_W={BODY_W:.3f} HEAD_W={HEAD_W:.3f}")
    prev = None
    for pi in range(d.page_count):
        raw = d[pi].get_text("dict")
        ys = {}
        for b in raw["blocks"]:
            for l in b.get("lines", []):
                for s in l.get("spans", []):
                    t = s["text"].strip()
                    if t: ys.setdefault(round(s["origin"][1], 1), []).append(t)
        keys = sorted(ys)
        tag = ys[keys[0]][0] if keys else "?"
        nlines = len(keys) - 1
        if not tag.startswith("A"): continue
        i = int(tag[1:]); ir = ARMS[i] * 10
        limit = CONTENT/20 - IND_L/20 - ir/20
        mark = ""
        if prev is not None and prev[1] != nlines: mark = "   <<< FLIP"
        print(f"  A{i:02d} ind_right={ir/20:6.2f}pt limit={limit:7.3f} lines={nlines}{mark}")
        if mark:
            # X window: between the two arms
            lo = CONTENT/20 - IND_L/20 - prev[0]/20 - BODY_W
            hi = CONTENT/20 - IND_L/20 - ir/20 - BODY_W
            print(f"       => X in ({hi:.3f}, {lo:.3f}]   (keep needs BODY_W <= limit - X)")
        prev = (ir, nlines)

if __name__ == "__main__":
    {"gen": gen, "measure": measure, "read": read}[sys.argv[1]]()
