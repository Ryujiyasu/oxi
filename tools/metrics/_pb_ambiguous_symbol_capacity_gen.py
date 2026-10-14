# -*- coding: utf-8 -*-
"""Does Word measure an East-Asian-AMBIGUOUS symbol drawn by a fallback face
(U+2610 BALLOT BOX in a Yu Mincho run -> Segoe UI Symbol, 0.86 em) at that
advance when breaking the line, or as a full cell?  The cell_pair_boundaries
fixture (9 full-width characters + the symbol, exact 280, jc both, cell 95 /
97.5) cannot tell: both readings wrap at 95 and fit at 97.5.  Sweep the cell
width in quarter points with the symbol and with a kanji control:
    w{95.5,96,96.5,97,97.25}_{box,kanji,chk,sq}   (sq U+25A1 is in Yu Mincho itself)
Read: whether the 10th character stays on line 1 (Word PDF; Oxi --dump-glyphs).

    python tools/metrics/_pb_ambiguous_symbol_capacity_gen.py gen
    python tools/metrics/_pb_ambiguous_symbol_capacity_gen.py word
    python tools/metrics/_pb_ambiguous_symbol_capacity_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_ambiguous_symbol_capacity"
BODY = "甲乙・丙丁・戊己・"  # 甲乙・丙丁・戊己・
LAST = {"box": "☐", "kanji": "庚", "chk": "✔", "sq": "□"}
WIDTHS = [95.5, 96.0, 96.5, 97.0, 97.25]
ARMS = {"w%s_%s" % (str(w).replace(".", "p"), k): (w, c) for w in WIDTHS for k, c in LAST.items()}
YU = "Yu Mincho"


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (w, last) in ARMS.items():
        tw = int(round(w * 20))
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
                  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        tbl = ('<w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>' % tw +
               '<w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="%d"/><w:tcMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tcMar></w:tcPr>' % tw +
               '<w:p><w:pPr><w:spacing w:after="0" w:line="280" w:lineRule="exact"/><w:jc w:val="both"/></w:pPr>'
               '<w:r><w:rPr><w:rFonts w:ascii="' + YU + '" w:hAnsi="' + YU + '" w:eastAsia="' + YU + '"/><w:sz w:val="21"/><w:lang w:val="ja-JP" w:eastAsia="ja-JP"/></w:rPr>'
               '<w:t>' + BODY + last + '</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body><w:p><w:r><w:t>BEFORE</w:t></w:r></w:p>' + tbl +
               '<w:p><w:r><w:t>AFTER</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / (arm + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
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
    for arm, (wd, last) in ARMS.items():
        pg = fitz.open(str(OUT / (arm + ".pdf")))[0]; rows = {}
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if s["text"].strip() and s["text"].strip() not in ("BEFORE", "AFTER"):
                        y = round(s["origin"][1], 1); rows.setdefault(y, []); rows[y].append((s["font"], round(s["bbox"][0], 1), round(s["bbox"][2], 1), s["text"]))
        lines = [("".join(x[3] for x in sorted(v, key=lambda x: x[1])), sorted({x[0] for x in v})) for y, v in sorted(rows.items())]
        res[arm] = lines
        print("%-14s word %s" % (arm, lines))
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / (arm + ".docx")), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        g = json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]; rows = {}
        for x in g:
            if x["char"].strip() and x["char"] not in "BEFORAFT":
                rows.setdefault(round(x["baseline"], 1), []).append((x["x"], x["char"]))
        lines = ["".join(c for _, c in sorted(v)) for y, v in sorted(rows.items())]
        print("%-14s oxi  %s" % (arm, lines))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
