# -*- coding: utf-8 -*-
"""Which glyphs does Word draw for the Japanese kana numbering formats?

tokyoshugyo's （エ）（オ）list is `aiueoFullWidth`, and Word renders it in
KATAKANA while Oxi renders hiragana (crates/oxidocs-core/src/parser/numbering.rs
uses あ/い/う for aiueo* and い/ろ/は for iroha*). Read all four formats straight
out of Word so the table is measured rather than recalled.

    python _pb_numfmt_gen.py gen
    python _pb_numfmt_gen.py pdf      # Word truth
    python _pb_numfmt_gen.py oxi      # Oxi, same arms
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_numfmt")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

FACE = "ＭＳ 明朝"
SZ_HP = 21
FORMATS = ["aiueo", "aiueoFullWidth", "iroha", "irohaFullWidth"]
# OXI_PB_ITEMS=48 walks the whole kana sequence (it spills onto extra pages, so
# _collect accumulates until the next F#START rather than reading one page).
ITEMS = int(os.environ.get("OXI_PB_ITEMS", "6"))


def docx():
    return os.path.join(OUT, "numfmt.docx")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for fi, fmt in enumerate(FORMATS):
        body.append('<w:p><w:pPr>%s<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"'
                    ' w:eastAsia="%s"/><w:sz w:val="%d"/></w:rPr></w:pPr>'
                    '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
                    '<w:sz w:val="%d"/></w:rPr><w:t>F%dSTART</w:t></w:r></w:p>'
                    % ("<w:pageBreakBefore/>" if fi else "", FACE, FACE, FACE, SZ_HP,
                       FACE, FACE, FACE, SZ_HP, fi))
        for k in range(ITEMS):
            body.append('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/>'
                        '<w:numId w:val="%d"/></w:numPr>'
                        '<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
                        '<w:sz w:val="%d"/></w:rPr></w:pPr>'
                        '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
                        '<w:sz w:val="%d"/></w:rPr><w:t>ITEM</w:t></w:r></w:p>'
                        % (fi + 1, FACE, FACE, FACE, SZ_HP, FACE, FACE, FACE, SZ_HP))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
           '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" '
           'w:header="851" w:footer="992" w:gutter="0"/>'
           "</w:sectPr></w:body></w:document>")
    numbering = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + NS + ">"]
    for fi, fmt in enumerate(FORMATS):
        numbering.append(
            '<w:abstractNum w:abstractNumId="%d"><w:multiLevelType w:val="singleLevel"/>'
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="%s"/>'
            '<w:lvlText w:val="(%%1)"/><w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="840" w:hanging="420"/></w:pPr></w:lvl></w:abstractNum>'
            % (fi, fmt))
    for fi in range(len(FORMATS)):
        numbering.append('<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>'
                         % (fi + 1, fi))
    numbering.append("</w:numbering>")
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:rPr><w:sz w:val="%d"/></w:rPr></w:style>'
              "</w:styles>" % (FACE, FACE, FACE, SZ_HP))
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/numbering.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
                    "</Types>")
    drels = DRELS.replace("</Relationships>",
                          '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/'
                          'officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
                          "</Relationships>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/numbering.xml", "".join(numbering))
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(FORMATS), "formats x", ITEMS, "items")


def report(per, who):
    print("== %s ==" % who)
    for fi, fmt in enumerate(FORMATS):
        marks = per.get(fi) or []
        print("%-16s %s" % (fmt, " ".join(marks) if marks else "MISSING"))


def _collect(pagetexts):
    per, cur = {}, None
    for txt in pagetexts:
        m = re.search(r"F(\d)START", txt)
        if m:
            cur = int(m.group(1))
            per.setdefault(cur, [])
        if cur is None:
            continue
        per[cur].extend(re.findall(r"[(（]\s*(\S)\s*[)）]", txt))
    return {k: v[:ITEMS] for k, v in per.items()}


def pdf():
    import fitz
    import win32com.client as w
    out = docx().replace(".docx", ".pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx(), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(out, 17)
    finally:
        d.Close(False)
        app.Quit()
    doc = fitz.open(out)
    report(_collect([doc[i].get_text() for i in range(doc.page_count)]), "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "numfmt_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "nf"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    texts = []
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        rows = {}
        for e in pg["elements"]:
            if e["type"] == "text":
                rows.setdefault(round(e["y"], 1), []).append((e["x"], e.get("text") or ""))
        texts.append("\n".join("".join(t for _, t in sorted(v))
                               for _, v in sorted(rows.items())))
    report(_collect(texts), "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
