# -*- coding: utf-8 -*-
"""Does a kinsoku-off paragraph (w:kinsoku val=0) still squeeze punctuation to pull a character in?

ja/creative__285f61cb (compat 14, compressPunctuation, every paragraph kinsoku=0): Word fills
the 20-cell line «人がいることが分かったからだ。もっとも、» and starts the next line with «「»;
Oxi halves «、» and «「» to set «「» as a 21st character.

Arms (one document each, fresh Word process): compat 15 / 14 x kinsoku on / off. Each
paragraph fills the 40-cell line exactly and then offers one more character; prints the
first line's last character and the second line's first two, Word and Oxi.

    python tools/metrics/_pb_kinsoku_off_squeeze_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_kinsoku_off_squeeze")
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
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="doNotCompress"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % W
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="20"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/><w:jc w:val="left"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') % W
CELLS = 40
RUN_LANG = os.environ.get("PB_RUN_LANG")
JC = os.environ.get("PB_JC", "both")
CSC = os.environ.get("PB_CSC", "compressPunctuation")
GRID = os.environ.get("PB_GRID")
IND_EXTRA = int(os.environ.get("PB_IND_EXTRA", "0"))
BALANCE = os.environ.get("PB_BALANCE") == "1"
LATIN_FONT = os.environ.get("PB_LATIN_FONT", "ＭＳ 明朝")  # ascii / hAnsi face  # <w:balanceSingleByteDoubleByteWidth/> (the S1585 probe had it)  # twips added to the right indent (sub-cell sweeps)  # docGrid type (lines / linesAndChars) with linePitch 360  # eastAsia lang on every run (en-US = no kinsoku at all)
TEXTS = [
    "あ、" * 20 + "いうえお",
    "あ" * 38 + "、、" + "いうえお",
    "あ" * 39 + "、" + "いうえお",
    "あ、" * 20 + "「うえお",
    "あ" * 39 + "、" + "「うえお",
    "あ、" * 20 + "」うえお",
    "あ" * 39 + "、" + "」うえお",
    "あ" * 39 + "。" + "）うえお",
    "あ、" * 20 + "。うえお",
    "あ、" * 20 + "・うえお",
    # 20-cell lines (ja/creative__285f61cb's «…からだ。もっとも、» + «「»)
    (20, "あ" * 18 + "。、" + "「うえお"),
    (20, "あ" * 18 + "。、" + "いうえお"),
    (20, "あ" * 16 + "。あ、、" + "「うえお"),
    (20, "あ" * 16 + "。あ、、" + "いうえお"),
]
if os.environ.get("PB_HALF") == "1":
    # a half-width «a» needs half a cell: do the marks' blanks lend it?
    TEXTS = [
        "あ、" * 20 + "aうえお",             # twenty 、 on the line
        "あ" * 39 + "、" + "aうえお",         # one 、
        "あ" * 37 + "「あ」" + "aうえお",     # one pair of brackets
        "あ" * 38 + "「あ" + "aうえお",       # one opening bracket
        "あ" * 40 + "aうえお",               # no mark at all
    ]
if os.environ.get("PB_HALF") == "2":
    # script and line length: half-width katakana vs Latin, 14-cell vs 40-cell lines
    TEXTS = [
        (40, "あ、" * 20 + "ｱうえお"),
        (40, "あ" * 39 + "、" + "ｱうえお"),
        (14, "あ" * 13 + "、" + "aうえお"),
        (14, "あ" * 13 + "、" + "ｱうえお"),
        (14, "あ" * 12 + "「あ」" + "aうえお"),
        (40, "あ" * 38 + "、" + "abうえお"),
        (40, "あ、" * 20 + "abうえお"),
    ]
if os.environ.get("PB_HALF") == "3":
    # a Latin unit after kana carries an autospace gap (0.25em): is the gap, not the mark, the credit?
    TEXTS = [
        (40, "あ" * 39 + "abうえお"),            # gap + 1.0  (over 0.25)
        (40, "あ" * 39 + "10うえお"),            # digits, same geometry
        (40, "あ" * 38 + "、" + "abうえお"),     # no gap after 、, fits exactly (control)
        (40, "あ" * 38 + "、い" + "abうえお"),   # gap + 、 (over 1.25)
        (40, "あ" * 39 + "abcうえお"),           # gap + 1.5 (over 0.75)
        (40, "あ" * 38 + "、" + "abcうえお"),    # 、 only (over 0.5)
        (40, "あ" * 38 + "い" + "aうえお"),      # gap + 0.5 (over -0.25: fits) control
        (40, "あ" * 39 + "aうえお"),             # gap + 0.5 (over 0.25... 39+0.25+0.5 = 39.75 fits) control
    ]
if os.environ.get("PB_HALF") == "6":
    TEXTS = [(40, "あ" * 39 + "abうえお")]   # need 0.25 cell + PB_IND_EXTRA/200 cells, one gap
if os.environ.get("PB_HALF") == "5":
    # how much of the autospace gaps does a justified compat-15 line squeeze to keep its last word?
    TEXTS = [
        (40, "あ" * 39 + "abうえお"),              # need 0.25 cell, 1 gap, mid line
        (40, "あ" * 39 + "ab"),                    # same, paragraph end (last line)
        (40, "あ" * 38 + "aあbcうえお"),           # need 0.25, 3 gaps
        (40, "あ" * 38 + "aあbcdうえお"),          # need 0.75, 3 gaps
        (40, "あ" * 37 + "aあbあcdうえお"),        # need 0.75 + ... : 37+0.5+1+0.5+1+1 = 41 (+5 gaps 1.25) wraps anyway? -> need 2.25
        (40, "あ" * 38 + "ab" + "あ" + "cうえお"),  # need 0.5 + 3 gaps 0.75: 38+1+1+0.5 = 40.5 -> over 1.25 with 3 gaps
    ]
if os.environ.get("PB_HALF") == "4":
    # need vs the slack a wrap would leave: half-width ｱ before the word moves the wrap point
    TEXTS = [
        (40, "あ" * 38 + "、ｱ" + "abうえお"),    # need 0.5, wrap leaves 0.5 (tie)
        (40, "あ" * 38 + "、ｱ" + "abcうえお"),   # need 1.0, wrap leaves 0.5
        (40, "あ" * 37 + "、ｱ" + "abcうえお"),   # need 0.5, wrap leaves 1.0
        (40, "あ" * 37 + "、ｱ" + "abcdうえお"),  # need 1.0, wrap leaves 1.0
        (40, "あ" * 38 + "ｱ" + "abcうえお"),     # fits exactly (control)
        (40, "あ" * 38 + "、" + "a bcうえお"),    # 'a' fits, 'bc' needs 1.0 after a space
    ]


def build(compat, koff):
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="%s"/>%s'
                '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>') % (W, CSC, '<w:balanceSingleByteDoubleByteWidth/>' if BALANCE else '', compat)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="20"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
              '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/><w:jc w:val="%s"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') % (W, LATIN_FONT, LATIN_FONT, JC)
    body = "".join('<w:p><w:pPr>%s<w:ind w:right="%d"/></w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p><w:p/>'
                   % ('<w:kinsoku w:val="0"/>' if koff else '', 9026 - (t[0] if isinstance(t, tuple) else CELLS) * 200 + IND_EXTRA,
                      '<w:rPr><w:lang w:eastAsia="%s"/></w:rPr>' % RUN_LANG if RUN_LANG else '',
                      t[1] if isinstance(t, tuple) else t) for t in TEXTS)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/>%s</w:sectPr></w:body></w:document>'
           % (W, body, '<w:docGrid w:type="%s" w:linePitch="360"/>' % GRID if GRID else ''))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "ks_c%d_%s%s%s.docx" % (compat, "off" if koff else "on", "_" + RUN_LANG if RUN_LANG else "", "_half" + os.environ.get("PB_HALF", "") + ("_x%d" % IND_EXTRA if IND_EXTRA else "") + ("_bal" if BALANCE else "") + ("_" + LATIN_FONT.replace(" ", "") if LATIN_FONT != "ＭＳ 明朝" else "") if os.environ.get("PB_HALF") else "") if JC == "both" and CSC == "compressPunctuation" and not GRID else "ks_c%d_%s_%s_%s_%s%s.docx" % (compat, "off" if koff else "on", RUN_LANG or "ja", JC, CSC, "_grid" + GRID if GRID else ""))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    return path


def pairs(lines):
    out = []
    i = 0
    lines = [l.replace(" ", "") for l in lines if l.replace(" ", "")]
    while i + 1 < len(lines):
        if lines[i].startswith("あ"):
            out.append((lines[i][-1], lines[i + 1][:2])); i += 2
        else:
            i += 1
    return out


def word_lines(pdf):
    import pymupdf
    res = []
    for pg in pymupdf.open(pdf):
        rows = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        rows.setdefault(round(c["origin"][1]), []).append((c["origin"][0], c["c"]))
        res += ["".join(c for _, c in sorted(rows[k])) for k in sorted(rows)]
    return res


def oxi_lines(exe, docx):
    import subprocess, json, tempfile
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    res = []
    for p in json.load(open(g, encoding="utf-8"))["pages"]:
        rows = {}
        for c in p["glyphs"]:
            rows.setdefault(round(c["baseline"]), []).append(c)
        res += ["".join(c["char"] for c in sorted(rows[k], key=lambda c: c["x"])) for k in sorted(rows)]
    return res


if __name__ == "__main__":
    import win32com.client
    for compat in (15, 14):
        for koff in (False, True):
            p = build(compat, koff); pdf = p[:-5] + ".pdf"
            w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
            try:
                d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
            finally:
                w.Quit()
            wp = pairs(word_lines(pdf))
            op = pairs(oxi_lines(os.environ["PB_OXI"], p)) if os.environ.get("PB_OXI") else [None] * len(wp)
            for t, a, b in zip(TEXTS, wp, op):
                t = t[1] if isinstance(t, tuple) else t
                print("c%d %-3s %-12s Word %-14s Oxi %-14s %s" % (compat, "off" if koff else "on", t[-6:], a, b, "" if a == b else "<<"), flush=True)
