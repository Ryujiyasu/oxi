# -*- coding: utf-8 -*-
"""S1615 minimal repro: floating table (vertAnchor=text) with tblpYSpec.

Self-authored, Calibri 11, no styles part beyond defaults.  Body:
  P1 "Line one" / P2 NOTE / floating table (1x1 "TCELL", 3 lines tall) / P3 "After"
Arms:
  bottom_1   tblpYSpec=bottom, NOTE one line
  bottom_2   tblpYSpec=bottom, NOTE two lines
  top_1      tblpYSpec=top
  center_1   tblpYSpec=center
  y0_1       tblpY=0 (no YSpec)
  inline_1   no tblpPr
  bottom_e   tblpYSpec=bottom, NOTE empty
Read (Word COM Information(6)): P2, the table cell, P3.

    python tools/metrics/_pb_ftbl_yspec_min_gen.py gen
    python tools/metrics/_pb_ftbl_yspec_min_gen.py word
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "tests/fixtures/ftbl_yspec_min"
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def p(t):
    return f'<w:p><w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>' if t else "<w:p/>"


def table(pos):
    tp = f'<w:tblpPr w:leftFromText="142" w:rightFromText="142" w:vertAnchor="text" w:horzAnchor="margin" {pos}/>' if pos else ""
    cell = "".join(f'<w:p><w:r><w:t>TCELL{i}</w:t></w:r></w:p>' for i in range(3))
    return (f'<w:tbl><w:tblPr>{tp}<w:tblW w:w="9000" w:type="dxa"/>'
            '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>'
            '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{cell}</w:tc></w:tr></w:tbl>')


LONG = "NOTE " + "word " * 30
ARMS = {
    "bottom_1": ('w:tblpYSpec="bottom"', "NOTE short"),
    "bottom_2": ('w:tblpYSpec="bottom"', LONG),
    "top_1": ('w:tblpYSpec="top"', "NOTE short"),
    "center_1": ('w:tblpYSpec="center"', "NOTE short"),
    "y0_1": ('w:tblpY="0"', "NOTE short"),
    "inline_1": (None, "NOTE short"),
    "bottom_e": ('w:tblpYSpec="bottom"', ""),
}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for name, (pos, note) in ARMS.items():
        body = p("Line one") + p(note) + table(pos) + p("After") + \
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
        doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}</w:body></w:document>'
        with zipfile.ZipFile(OUT / f"{name}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
            z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
            z.writestr("word/document.xml", doc)
    print("ok")


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    res = {}
    try:
        for name in ARMS:
            tmp = os.path.join(tempfile.mkdtemp(), name + ".docx")
            shutil.copy(OUT / f"{name}.docx", tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
            try:
                ys = []
                for i in range(1, d.Paragraphs.Count + 1):
                    r = d.Paragraphs(i).Range
                    ys.append((r.Text[:8].strip(), d.Range(r.Start, r.Start).Information(6)))
                res[name] = ys
            finally:
                d.Close(0)
            print(name, ys, flush=True)
    finally:
        w.Quit()
    (OUT / "word.json").write_text(json.dumps(res, indent=1), encoding="utf-8")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word}[sys.argv[1]]()
