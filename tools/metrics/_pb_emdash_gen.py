# -*- coding: utf-8 -*-
"""Derive Word's line-break opportunity around an EM DASH (U+2014).

UAX #14 puts U+2014 in class B2 (break allowed on BOTH sides).  Oxi has no
break opportunity there at all, so an NBSP-joined run like
`1300 mm —400 mm` becomes ONE unbreakable token
(policies__00148f8d p68 (c): Word keeps `... not over 1300 mm —` on the
line and starts the next with `400 mm ...`; Oxi sends the whole token down).

Each arm is one section = one page: a paragraph whose right indent is swept so
that the dash lands at different distances from the margin.  Reading the Word
PDF tells us which side Word actually breaks on.

  python tools/metrics/_pb_emdash_gen.py gen
  python tools/metrics/_pb_emdash_gen.py measure     (Word COM -> PDF)
  python tools/metrics/_pb_emdash_gen.py read
"""
import os
import subprocess
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = Path("pipeline_data/_pb_emdash")
DOCX = OUT / "emdash.docx"
PDF = OUT / "emdash.pdf"

# 4 shapes x right-indent sweep.  The filler word count is fixed; only the
# right indent moves, so the dash walks across the margin.
SHAPES = {
    # NBSP-joined, exactly the target's shape
    "NB": "wwww wwww wwww not over 1300 mm —400 mm from the centre",
    # ordinary spaces around the dash
    "SP": "wwww wwww wwww not over 1300 mm — 400 mm from the centre",
    # dash glued to both neighbours (no space at all)
    "GL": "wwww wwww wwww not over 1300mm—400mm from the centre",
    # en dash control (U+2013, same ambiguous class)
    "EN": "wwww wwww wwww not over 1300 mm –400 mm from the centre",
}
RIGHTS = list(range(3000, 4801, 100))  # twips: sweeps the dash across the margin


def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    body = []
    for shape, text in SHAPES.items():
        for r in RIGHTS:
            tag = f"{shape}{r:04d}"
            body.append(
                '<w:p><w:pPr>'
                f'<w:ind w:left="0" w:right="{r}"/>'
                '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
                '</w:pPr>'
                f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
                f'<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{tag} {esc(text)}</w:t></w:r>'
                '</w:p>'
            )
            body.append(
                '<w:p><w:pPr><w:sectPr>'
                '<w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
                ' w:header="720" w:footer="720" w:gutter="0"/>'
                '</w:sectPr></w:pPr></w:p>'
            )
    doc = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>' + "".join(body) +
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
        '</w:body></w:document>'
    )
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
    print("wrote", DOCX, "arms", len(SHAPES) * len(RIGHTS))


def measure():
    from win32com.client import DispatchEx
    word = None
    try:
        word = DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        d = word.Documents.Open(str(DOCX.resolve()), ReadOnly=True,
                                AddToRecentFiles=False)
        d.ExportAsFixedFormat(str(PDF.resolve()), 17)
        n = d.ComputeStatistics(2)
        d.Close(False)
        print("exported", PDF, "pages", n)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def read():
    import fitz
    d = fitz.open(str(PDF))
    print(f"{'arm':8s} {'lines':5s}  line1 tail")
    for pi in range(d.page_count):
        rows = []
        for b in d[pi].get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if not s["text"].strip():
                        continue
                    rows.append([round(s["origin"][1], 2), round(s["bbox"][2], 2),
                                 s["text"]])
        rows.sort()
        m = []
        for r in rows:
            if m and abs(m[-1][0] - r[0]) <= 0.75:
                m[-1][2] += r[2]
                m[-1][1] = r[1]
            else:
                m.append(r)
        if not m:
            continue
        tag = m[0][2].split()[0]
        tail = m[0][2].rstrip()[-16:]
        print(f"{tag:8s} {len(m):5d}  x1={m[0][1]:7.2f}  ...{tail!r}")


if __name__ == "__main__":
    {"gen": gen, "measure": measure, "read": read}[sys.argv[1]]()
