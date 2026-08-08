# -*- coding: utf-8 -*-
"""VML float wrap probe: does the text AFTER a wrapping VML float sit BESIDE
it (in the free lane) or BELOW it?  Sweeps the float width so the free lanes
shrink, on both `square` and `tight` wrap modes, with a TEXT follower and an
EMPTY-paragraph follower.

  python tools/metrics/_pb_vmlwrap_gen.py gen
  python tools/metrics/_pb_vmlwrap_gen.py measure
  python tools/metrics/_pb_vmlwrap_gen.py read
"""
import sys, zipfile, shutil, subprocess, os
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = Path("pipeline_data/_pb_vmlwrap"); OUT.mkdir(parents=True, exist_ok=True)
DOCX = OUT / "vmlwrap.docx"; PDF = OUT / "vmlwrap.pdf"

# A4 595.32 x 841.92, margins 36 -> content [36, 559.32], width 523.32
WIDTHS = [120.0, 240.0, 340.0, 400.0, 411.5, 440.0, 470.0, 500.0]
WRAPS  = ["square", "tight"]
FOLLOW = ["text", "empty"]

def para(txt, sz=22):
    return ('<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr><w:sz w:val="{sz}"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:sz w:val="{sz}"/></w:rPr><w:t xml:space="preserve">{txt}</w:t></w:r></w:p>')

def empty(sz=22):
    return ('<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr><w:sz w:val="{sz}"/></w:rPr></w:pPr></w:p>')

def anchor(tag, w, wrap):
    # float 88pt tall, 22.8pt below the anchor line, left offset 32pt
    return ('<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:pict>'
            f'<v:shape id="s{tag}" type="#_x0000_t136" '
            f'style="position:absolute;margin-left:32pt;margin-top:22.8pt;'
            f'width:{w}pt;height:88pt;z-index:-251658224" fillcolor="#ffc000">'
            f'<v:textpath string="X"/><w10:wrap type="{wrap}"/>'
            '</v:shape></w:pict></w:r>'
            f'<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>ANCHOR{tag}</w:t></w:r></w:p>')

SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
        'w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>')

def gen():
    body = []
    tag = 0
    labels = []
    for wrap in WRAPS:
        for fol in FOLLOW:
            for w in WIDTHS:
                tag += 1
                lab = f"{wrap[0].upper()}{fol[0].upper()}{tag:02d}"
                labels.append((lab, wrap, fol, w))
                body.append(para(f"TOP{lab}"))
                body.append(anchor(lab, w, wrap))
                if fol == "empty":
                    body.append(empty()); body.append(empty())
                else:
                    body.append(para(f"MID{lab}"))
                body.append(para(f"END{lab}"))
                # one section per arm so every arm starts at the page top
                body.append('<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                            '<w:pgMar w:top="720" w:right="720" w:bottom="720" '
                            'w:left="720" w:header="0" w:footer="0" w:gutter="0"/>'
                            '</w:sectPr></w:pPr></w:p>')
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:v="urn:schemas-microsoft-com:vml" '
           'xmlns:w10="urn:schemas-microsoft-com:office:word" '
           'xmlns:o="urn:schemas-microsoft-com:office:office">'
           '<w:body>' + "".join(body) + SECT + '</w:body></w:document>')
    base = Path("pipeline_data/docx_corpus/en/correspondence/000f9471dada0bc4.docx")
    shutil.copy(base, DOCX)
    zin = zipfile.ZipFile(base)
    items = [i for i in zin.infolist() if i.filename != "word/document.xml"]
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        for i in items:
            z.writestr(i, zin.read(i.filename))
        z.writestr("word/document.xml", xml)
    (OUT / "labels.txt").write_text("\n".join(f"{l}\t{wr}\t{f}\t{w}" for l, wr, f, w in labels), encoding="utf-8")
    print(f"wrote {DOCX} ({len(labels)} arms)")

def measure():
    import win32com.client as win32
    app = win32.DispatchEx("Word.Application"); app.Visible = False
    try:
        d = app.Documents.Open(str(DOCX.resolve()), ReadOnly=True)
        d.ExportAsFixedFormat(str(PDF.resolve()), 17)
        d.Close(False)
    finally:
        app.Quit()
    print("exported", PDF)

def read():
    import fitz
    labels = [l.split("\t") for l in (OUT / "labels.txt").read_text(encoding="utf-8").splitlines()]
    doc = fitz.open(PDF)
    pos = {}
    for pi in range(doc.page_count):
        for b in doc[pi].get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                for s in ln["spans"]:
                    t = s["text"].strip()
                    if t.startswith(("TOP", "MID", "END", "ANCHOR")):
                        pos[t] = (pi, round(s["origin"][0], 2), round(s["origin"][1], 2))
    print(f"{'lab':>8} {'wrap':>7} {'fol':>6} {'w':>6} {'lanes(L,R)':>14} "
          f"{'anchorY':>8} {'nextX':>8} {'nextY':>8}  verdict")
    for lab, wrap, fol, w in labels:
        w = float(w)
        L = 32.0; R = 523.32 - (32.0 + w)
        a = pos.get(f"ANCHOR{lab}")
        n = pos.get(f"MID{lab}") or pos.get(f"END{lab}")
        if not a or not n:
            print(f"{lab:>8} {wrap:>7} {fol:>6} {w:6.1f}  (missing)"); continue
        # band bottom (relative to the anchor baseline) = 22.8 + 88
        below = n[2] - a[2] > 100.0 or n[0] > 3.0 + 36.0 + 0.0 and False
        verdict = "BELOW" if (n[2] - a[2]) > 100.0 else "beside"
        print(f"{lab:>8} {wrap:>7} {fol:>6} {w:6.1f} ({L:5.1f},{R:6.1f}) "
              f"{a[2]:8.2f} {n[1]:8.2f} {n[2]:8.2f}  {verdict}")

if __name__ == "__main__":
    {"gen": gen, "measure": measure, "read": read}[sys.argv[1]]()
