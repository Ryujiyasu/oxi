# -*- coding: utf-8 -*-
"""Baseline of a Symbol-bullet list line (U+F0B7 marker from the numbering
level's Symbol face) against a plain paragraph of the same font.  blind-G
en/correspondence__00afe08d and en/creative__03aea325 (no-type docGrid 360,
Century Gothic / Calibri 11pt, line 276 / 240 auto): the first bullet line sits
1.0pt higher in Oxi than in Word while plain lines match; the Latin baseline
law (S1047) excludes Symbol-bullet lines and leaves them to the older
centring.  Arms: three paragraphs (plain «Heading text:» / bullet «Assist with
all major ministries» / bullet «Create and set up welcoming environments»)
then «end»:
    cg276_g   Century Gothic 11pt, line 276, no-type grid 360   (00afe08d)
    cal240_g  Calibri 11pt, line 240, grid 360                   (03aea325)
    cg276_n   Century Gothic, line 276, no docGrid
    cal259_g  Calibri, line 259, grid 360
    cg276_g14 Century Gothic 14pt, line 276, grid 360
Read: baselines of the heading, the bullet glyph and the bullet text (Word PDF
spans / Oxi glyphs).

    python tools/metrics/_pb_symbol_bullet_base_gen.py gen
    python tools/metrics/_pb_symbol_bullet_base_gen.py word
    python tools/metrics/_pb_symbol_bullet_base_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_symbol_bullet_base"
# (font, sz half-points, line, grid)
ARMS = {"cg276_g": ("Century Gothic", 22, 276, True), "cal240_g": ("Calibri", 22, 240, True),
        "cg276_n": ("Century Gothic", 22, 276, False), "cal259_g": ("Calibri", 22, 259, True),
        "cg276_g14": ("Century Gothic", 28, 276, True)}
NUMBERING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + T.NS + '>'
             '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
             '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/>'
             '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>'
             '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (font, sz, line, grid) in ARMS.items():
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/><w:lang w:val="en-US" w:eastAsia="en-US"/>' % (font, font, font, font, sz, sz) +
                  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="%d" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' % line +
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        p = lambda t: '<w:p><w:r><w:t xml:space="preserve">' + t + '</w:t></w:r></w:p>'
        b = lambda t: '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">' + t + '</w:t></w:r></w:p>'
        dg = '<w:docGrid w:linePitch="360"/>' if grid else ''
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + p("Heading text:") + b("Assist with all major ministries") + b("Create and set up welcoming environments") + p("end") +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' + dg + '</w:sectPr></w:body></w:document>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                          '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                                '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / (arm + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings); z.writestr("word/numbering.xml", NUMBERING); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def word():
    import win32com.client, fitz
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / (arm + ".docx"), t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / (arm + ".pdf")).resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm in ARMS:
        pg = fitz.open(str(OUT / (arm + ".pdf")))[0]; rows = []
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if s["text"].strip(): rows.append((round(s["origin"][1], 2), round(s["origin"][0], 1), s["font"], s["text"][:10]))
        rows.sort(); res[arm] = rows; print("%-10s word %s" % (arm, rows))
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / (arm + ".docx")), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        g = json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]; rows = {}
        for x in g:
            if x["char"].strip():
                k = (round(x["baseline"], 2), x["font_family"]); rows.setdefault(k, [x["x"], ""]); rows[k][1] = (rows[k][1] + x["char"])[:10]
        print("%-10s oxi  %s" % (arm, [(k[0], round(v[0], 1), k[1], v[1]) for k, v in sorted(rows.items())]))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
