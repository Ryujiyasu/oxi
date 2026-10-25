# -*- coding: utf-8 -*-
"""The single-line row pitch of a table whose cells hold a ＭＳ face, in a compat-15
document WITHOUT a docGrid.

blind-G ja/legal__0e9cf147 (compat 15, no w:docGrid, docDefaults lang eastAsia
en-US, cells in ＭＳ ゴシック at the 11pt default): Word's rules step 14.76 =
14.28 (11 x the 83/64 win box, unrounded) + 0.48; Oxi's rows are 15.0 = 14.5 +
0.5 (`word_line_height_table_cell`: 15px -> 11.25 x 83/64 -> 14.59 -> 1/8 floor),
and the three notes below the table land 1.1-1.2 low.

Sheet: «前», a one-column table (tblBorders sz 4, full width) of ROWS rows of
«行» in FACE at SIZE, «後». Pitch = (last rule - first rule) / ROWS from the
Word PDF drawings (Oxi: --dump-layout border elements).

Arms: face (msgo ＭＳ ゴシック / msmin ＭＳ 明朝) x size (9 / 10.5 / 11 / 12 / 12.5) x docDefaults
eastAsia language (ja / en = en-US, the document's) x grid (none / lines 360).

    python tools/metrics/_pb_cell_nogrid_lh_gen.py        (PB_OXI=<exe>[;<exe>] adds Oxi)
"""
import json, os, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_cell_nogrid_lh"
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % W
FACES = {"msgo": "ＭＳ ゴシック", "msmin": "ＭＳ 明朝"}
ROWS = 8


def build(name, face, size, lang, grid):
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="22"/>'
              '<w:lang w:val="en-US" w:eastAsia="%s" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:widowControl w:val="0"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>') % (W, "ja-JP" if lang == "ja" else "en-US")
    rpr = '<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/><w:sz w:val="%d"/><w:lang w:eastAsia="ja-JP"/></w:rPr>' % (face, face, face, round(size * 2))
    rows = "".join('<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr><w:p><w:pPr>%s</w:pPr><w:r>%s<w:t>行%d</w:t></w:r></w:p></w:tc></w:tr>'
                   % (rpr, rpr, i) for i in range(ROWS))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
           + "".join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="auto"/>' % e for e in ("top", "left", "bottom", "right", "insideH", "insideV"))
           + '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' + rows + '</w:tbl>')
    g = '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid == "lines" else ""
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body><w:p><w:r><w:t>前</w:t></w:r></w:p>%s'
           '<w:p><w:r><w:t>後</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11910" w:h="16840"/>'
           '<w:pgMar w:top="1340" w:right="920" w:bottom="700" w:left="1020" w:header="0" w:footer="503"/>%s</w:sectPr></w:body></w:document>') % (W, tbl, g)
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{name}.docx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles); z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    return p


def pitch_word(pdf):
    import pymupdf
    ys = sorted({round(d["rect"].y0, 2) for d in pymupdf.open(pdf)[0].get_drawings() if d["rect"].width > 200 and d["rect"].height < 2})
    return (ys[-1] - ys[0]) / ROWS if len(ys) > 1 else float("nan")


def pitch_oxi(exe, docx):
    t = tempfile.mkdtemp(); lay = os.path.join(t, "l.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-layout=" + lay], capture_output=True)
    ys = sorted({round(e["y"], 2) for e in json.load(open(lay, encoding="utf-8"))["pages"][0]["elements"]
                 if e.get("type") == "border" and e.get("w", 0) > 200})
    return (ys[-1] - ys[0]) / ROWS if len(ys) > 1 else float("nan")


if __name__ == "__main__":
    import win32com.client
    exes = [e for e in os.environ.get("PB_OXI", "").split(";") if e]
    for grid in ("none", "lines"):
        for lang in ("en", "ja"):
            for f in FACES:
                for size in (9, 10.5, 11, 12, 12.5):
                    name = f"{f}_{size}_{lang}_{grid}"
                    p = build(name, FACES[f], size, lang, grid); pdf = p.with_suffix(".pdf")
                    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
                    try:
                        d = w.Documents.Open(str(p), ReadOnly=True); d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
                    finally:
                        w.Quit()
                    print(f"{name:22} W {pitch_word(pdf):.3f}" + "".join(f"  O {pitch_oxi(e, p):.3f}" for e in exes), flush=True)
