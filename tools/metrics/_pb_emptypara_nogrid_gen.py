# -*- coding: utf-8 -*-
"""The line box of an EMPTY paragraph when the section has no docGrid.  blind-G
ja/creative__6fd5a307 (bare package: document.xml + a theme whose minor East
Asian face is blank, Normal.dotm defaults after=160 line=278): between two 18pt
Hiragino Mincho Pro paragraphs (drawn in ＭＳ 明朝) sits a paragraph with an
EMPTY 18pt run; Word's gap author→body is 58.92 = (27.04 + 8) + (15.88 + 8),
i.e. the empty line is the 10.5pt ＭＳ 明朝 box × 278/240 (15.77), not the
18pt run and not the theme's Latin box (Times 12.07 × 1.158 = 14.0 = Oxi).
Arms (A4, margins 1440, explicit pPrDefault after=160 line=278, no docGrid;
paragraphs: «雪虫» 18pt ＭＳ 明朝 / the empty paragraph / «本文…» 18pt):
    run_empty      the empty paragraph holds an empty run with sz=36 (the document)
    run_none       the empty paragraph has no run
    mark_sz36      no run, paragraph-mark rPr sz=36
    mark_ms        no run, mark rPr rFonts eastAsia ＭＳ 明朝 (defaults otherwise)
    mark_yu        no run, mark rPr rFonts ascii Century eastAsia 游明朝
    mark_cal       no run, mark rPr rFonts ascii/eastAsia Calibri
    def_times_ms   docDefaults rFonts ascii Times New Roman eastAsia ＭＳ 明朝 (no mark rPr)
    def_times_cal  docDefaults rFonts ascii Times New Roman eastAsia Calibri
Read: baseline of «雪虫» and of the body line (Word PDF; Oxi --dump-glyphs); the
empty line = body − title − 27.04 − 16 (278/240 × 18 × 83/64 + two after=8).

    python tools/metrics/_pb_emptypara_nogrid_gen.py gen
    python tools/metrics/_pb_emptypara_nogrid_gen.py word
    python tools/metrics/_pb_emptypara_nogrid_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_emptypara_nogrid"
RUN18 = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="36"/></w:rPr>'
# (empty paragraph xml, docDefaults rFonts)
DEF = '<w:rFonts w:ascii="{a}" w:hAnsi="{a}" w:eastAsia="{e}" w:cs="{a}"/>'
ARMS = {"run_empty": ('<w:p><w:r>' + RUN18 + '<w:t></w:t></w:r></w:p>', ""),
        "run_none": ('<w:p/>', ""),
        "mark_sz36": ('<w:p><w:pPr><w:rPr><w:sz w:val="36"/></w:rPr></w:pPr></w:p>', ""),
        "mark_ms": ('<w:p><w:pPr><w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝"/></w:rPr></w:pPr></w:p>', ""),
        "mark_yu": ('<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="游明朝" w:hAnsi="Century"/></w:rPr></w:pPr></w:p>', ""),
        "mark_cal": ('<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri"/></w:rPr></w:pPr></w:p>', ""),
        "def_times_ms": ('<w:p/>', DEF.format(a="Times New Roman", e="ＭＳ 明朝")),
        "def_times_cal": ('<w:p/>', DEF.format(a="Times New Roman", e="Calibri")),
        # the document's shape: a theme part whose minor East Asian face is BLANK, fonts taken from the theme
        "th_blank_ea": ('<w:p/>', '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>'),
        "th_ms_ea": ('<w:p/>', '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>'),
        "th_blank_ea_run": ('<w:p><w:r><w:rPr><w:rFonts w:ascii="Hiragino Mincho Pro" w:hAnsi="Hiragino Mincho Pro"/><w:sz w:val="36"/></w:rPr><w:t></w:t></w:r></w:p>',
                            '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>'),
        "bare_blank_ea": ('<w:p><w:r><w:rPr><w:rFonts w:ascii="Hiragino Mincho Pro" w:hAnsi="Hiragino Mincho Pro"/><w:sz w:val="36"/></w:rPr><w:t></w:t></w:r></w:p>', None)}
THEME_SRC = REPO / "pipeline_data" / "docx_corpus" / "ja" / "creative" / "6fd5a3073c03a7b9.docx"


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    theme_blank = zipfile.ZipFile(THEME_SRC).read("word/theme/theme1.xml").decode("utf-8")
    theme_ms = theme_blank.replace('<a:ea typeface=""/>', '<a:ea typeface="ＭＳ 明朝"/>')
    for arm, (empty, deffonts) in ARMS.items():
        bare = deffonts is None
        deffonts = deffonts or ""
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  + deffonts + '<w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
                  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>'
               '<w:p><w:r>' + RUN18 + '<w:t>雪虫</w:t></w:r></w:p>' + empty +
               '<w:p><w:r>' + RUN18 + '<w:t>本文の一行目です。</w:t></w:r></w:p>'
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        themed = arm.startswith("th_") or bare
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                          + ('<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' if themed else '') + '</Types>')
        if bare:
            ct = ct.replace('<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>', '')
            ct = ct.replace('<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>', '')
        drels = T.DRELS.replace("</Relationships>", ('' if bare else '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>')
                                + ('<Relationship Id="rIdTh" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' if themed else '') + '</Relationships>')
        if bare:
            drels = drels.replace('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>', '')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", doc)
            if not bare:
                z.writestr("word/styles.xml", styles); z.writestr("word/settings.xml", settings)
            if themed:
                z.writestr("word/theme/theme1.xml", theme_ms if arm == "th_ms_ea" else theme_blank)
    print("ok", len(ARMS))


def report(arm, who, ys):
    gap = ys[1] - ys[0] - 27.04 - 16.0 if len(ys) >= 2 else None
    print(f"{arm:14s} {who} title={ys[0]:7.2f} body={ys[1]:7.2f}  empty line = {gap:6.2f}" if gap is not None else f"{arm:14s} {who} {ys}")
    return gap


def word():
    import win32com.client, fitz
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm in ARMS:
        pg = fitz.open(str(OUT / f"{arm}.pdf"))[0]; ys = []
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if s["text"].strip():
                        ys.append(round(s["origin"][1], 2))
        res[arm] = report(arm, "word", sorted(set(ys)))
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    for arm in ARMS:
        t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True, timeout=600)
        g = json.load(open(out, encoding="utf-8"))["pages"][0]["glyphs"]
        ys = sorted({round(x["baseline"], 2) for x in g if x["char"].strip()})
        report(arm, "oxi ", ys)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
