# -*- coding: utf-8 -*-
"""Does a superscript run grow its line, and by how much?

blind-G ja/educational__20aa108d: the footnote line «¹ 本稿における…» (9pt ＭＳ 明朝,
snapToGrid 0, lines 360 grid, the reference run in the «footnote reference»
character style = vertAlign superscript) sits 1.19 lower in Oxi; with the
superscript removed (`_pb_fn_snap0_gen.py` nosup) Oxi lands on Word's baseline.
So Oxi grows the line for the raised reference and Word does not -- here.

Sheet: N body paragraphs «X本文の行» where X is a superscript «1» (arm sup) or
absent (arm plain), then «後». The per-line pitch is (last - first) / (N-1) and the
first baseline is read against the top margin, so line growth shows in both.

Arms: face ＭＳ 明朝 x size 9 / 10.5 / 14 x grid (none / lines 360) x snapToGrid
(1 / 0) x sup (plain / sup / sup14: a 14pt superscript in a size-SIZE line / supc, supc14: the
superscript «1» in Century, as a footnote reference under a Latin minorHAnsi theme is).
PB_SUP=plain,supc,supc14 picks the sup arms; supcs = supc + a normal-size Century space run (the
footnote line's shape), cs = that space run alone.

    python tools/metrics/_pb_sup_line_gen.py          (PB_OXI=<exe>[;<exe>] adds Oxi)
"""
import json, os, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_sup_line"
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
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/><w:kern w:val="2"/><w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>') % W
N = 8


def build(name, size, grid, snap, sup):
    sz = round(size * 2)
    snap_ppr = "" if snap else '<w:snapToGrid w:val="0"/>'
    rpr = '<w:rPr><w:sz w:val="%d"/></w:rPr>' % sz
    if sup == "plain":
        lead = ""
    elif sup == "cs":  # a normal-size Century space run only (no superscript)
        lead = '<w:r><w:rPr><w:rFonts w:ascii="Century" w:hAnsi="Century"/><w:sz w:val="%d"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>' % sz
    else:
        ssz = 28 if sup.endswith("14") else sz
        face = '<w:rFonts w:ascii="Century" w:hAnsi="Century"/>' if sup.startswith("supc") else ""
        lead = '<w:r><w:rPr>%s<w:sz w:val="%d"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r>' % (face, ssz)
        if sup == "supcs":  # + a normal-size Century space run, as the footnote line has
            lead += '<w:r><w:rPr><w:rFonts w:ascii="Century" w:hAnsi="Century"/><w:sz w:val="%d"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>' % sz
    paras = "".join('<w:p><w:pPr>%s%s</w:pPr>%s<w:r>%s<w:t>本文の行%d</w:t></w:r></w:p>' % (snap_ppr, rpr, lead, rpr, i) for i in range(N))
    g = '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid == "lines" else ""
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s<w:p><w:r><w:t>後</w:t></w:r></w:p>'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>%s</w:sectPr></w:body></w:document>') % (W, paras, g)
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{name}.docx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES); z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    return p


def baselines_word(pdf):
    import pymupdf
    ys = []
    for b in pymupdf.open(pdf)[0].get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"] == "本":
                        ys.append(round(c["origin"][1], 2))
    return sorted(set(ys))


def baselines_oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    return sorted({round(c["baseline"], 2) for c in json.load(open(g, encoding="utf-8"))["pages"][0]["glyphs"] if c["char"] == "本"})


def fmt(ys):
    if len(ys) < 2:
        return "?"
    return "first %.2f pitch %.3f" % (ys[0], (ys[-1] - ys[0]) / (len(ys) - 1))


if __name__ == "__main__":
    import win32com.client
    exes = [e for e in os.environ.get("PB_OXI", "").split(";") if e]
    for grid in ("none", "lines"):
        for snap in (1, 0):
            if grid == "none" and snap == 0:
                continue
            for size in (9, 10.5, 14):
                for sup in os.environ.get("PB_SUP", "plain,sup,sup14").split(","):
                    if sup.endswith("14") and size == 14:
                        continue
                    name = f"{size}_{grid}_s{snap}_{sup}"
                    p = build(name, size, grid, snap, sup); pdf = p.with_suffix(".pdf")
                    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
                    try:
                        d = w.Documents.Open(str(p), ReadOnly=True); d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
                    finally:
                        w.Quit()
                    print(f"{name:22} W {fmt(baselines_word(pdf))}" + "".join(f"   O {fmt(baselines_oxi(e, p))}" for e in exes), flush=True)
