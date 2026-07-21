# -*- coding: utf-8 -*-
"""How many lines does a run of <w:br/> make inside a table cell?

A cell paragraph splits on soft breaks, and Oxi keeps the empty sublines but
gives them zero height because they hold no fragments. REPORT_K found the one
outlier among policies__0028d1be's 17 auto-spacing boundaries is exactly this:
`Physical Demands` opens with a <w:br/>, and Word's gap there is 41.67 against
27.8 elsewhere — one whole 13.8pt line more.

Before fixing it, the census says the shape is not rare: 60 leading, 12 trailing
and 16 consecutive empty sublines across 7 documents. So this measures all of
them at once, against a control with no break, using marker rows above and below
so the readout is baseline-to-baseline in one font.

  python _pb_cellbr_gen.py gen / measure / read
"""
from __future__ import annotations

import glob
import os
import re
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_cellbr"
HOST = REPO / "pipeline_data" / "docx_corpus" / "en" / "policies" / "0028d1bea47058b2.docx"

RPR = ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
       '<w:sz w:val="24"/></w:rPr>')
PPR = ('<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
       + RPR + '</w:pPr>')

# id -> the run content of the middle paragraph (markers bracket it)
CASES = {
    "CTRL":  '<w:r>{r}<w:t>MID</w:t></w:r>',
    "LEAD":  '<w:r>{r}<w:br/></w:r><w:r>{r}<w:t>MID</w:t></w:r>',
    "TRAIL": '<w:r>{r}<w:t>MID</w:t><w:br/></w:r>',
    "CONS2": '<w:r>{r}<w:br/><w:br/></w:r><w:r>{r}<w:t>MID</w:t></w:r>',
    "CONS3": '<w:r>{r}<w:br/><w:br/><w:br/></w:r><w:r>{r}<w:t>MID</w:t></w:r>',
    "MID2":  '<w:r>{r}<w:t>A</w:t><w:br/><w:br/><w:t>MID</w:t></w:r>',
}


def para(inner: str) -> str:
    return f'<w:p>{PPR}{inner.format(r=RPR)}</w:p>'


def cell(inner: str) -> str:
    return ('<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/>'
            '<w:tcBorders><w:top w:val="single" w:sz="8" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="8" w:color="000000"/></w:tcBorders>'
            '</w:tcPr>' + inner + '</w:tc>')


def case_table(cid: str) -> str:
    body = (para(f'<w:r>{{r}}<w:t>TOP{cid}</w:t></w:r>')
            + para(CASES[cid])
            + para(f'<w:r>{{r}}<w:t>BOT{cid}</w:t></w:r>'))
    return ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>'
            '<w:tblLayout w:type="fixed"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
            '<w:tr>' + cell(body) + '</w:tr></w:tbl>')


def gen() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    src = zipfile.ZipFile(HOST)
    doc = src.read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    body = "".join(case_table(c) + para('<w:r>{r}<w:t>-</w:t></w:r>') for c in CASES)
    with zipfile.ZipFile(OUT / "cellbr.docx", "w", zipfile.ZIP_DEFLATED) as z:
        for n in src.namelist():
            if n == "word/document.xml":
                z.writestr(n, head + body + (sect.group(0) if sect else "")
                           + "</w:body></w:document>")
            else:
                z.writestr(n, src.read(n))
    print(f"gen: {len(CASES)} cases in one document")


def measure() -> None:
    import win32com.client as win32
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        for path in sorted(glob.glob(str(OUT / "*.docx"))):
            pdf = path[:-5] + ".pdf"
            if os.path.exists(pdf):
                continue
            d = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
            try:
                d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf), ExportFormat=17)
            finally:
                d.Close(False)
            print("measured", os.path.basename(path), flush=True)
    finally:
        word.Quit()


def read() -> None:
    import fitz
    pdf = fitz.open(OUT / "cellbr.pdf")
    at = {}
    for page in pdf:
        for b in page.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if t.startswith(("TOP", "BOT")):
                    at.setdefault(t, round(ln["spans"][0]["origin"][1], 3))
    base = at.get("BOTCTRL", 0) - at.get("TOPCTRL", 0)
    print(f"  CTRL span (no break) = {base:.3f}")
    for cid in CASES:
        if cid == "CTRL":
            continue
        span = at.get(f"BOT{cid}", 0) - at.get(f"TOP{cid}", 0)
        extra = span - base
        print(f"  {cid:<6} span {span:8.3f}   extra {extra:+8.3f}   "
              f"= {extra / 13.8:.2f} lines of 13.8")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    {"gen": gen, "measure": measure, "read": read}[cmd]()
