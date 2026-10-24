# -*- coding: utf-8 -*-
"""Are the half-width katakana punctuation marks (U+FF61-FF65, ﾞﾟ) under Word's default kinsoku?

ja/technical__898a80c8: Word wraps before «｢» (U+FF62) where Oxi leaves it at the line end.

Each paragraph: N full-width «あ» then the mark under test then «いうえお»; N chosen so the
mark falls exactly at the line end (line-end test, opening marks) or at the next line's
start (line-start test, closing marks). Prints, per mark, the first line's last character
and the second line's first character, for Word and Oxi.

    python tools/metrics/_pb_halfwidth_kinsoku_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_halfwidth_kinsoku")
COMPAT = int(os.environ.get("PB_COMPAT", "15"))
JC = os.environ.get("PB_JC", "left")
EA_LANG = os.environ.get("PB_EA_LANG", "ja-JP")
ASCII_FONT = os.environ.get("PB_ASCII_FONT", "ＭＳ 明朝")
RUN_LANG = os.environ.get("PB_RUN_LANG")  # eastAsia lang on the run itself
DD_EA_FACE = os.environ.get("PB_DD_EA_FACE", "ＭＳ 明朝")
RUN_EA_FACE = os.environ.get("PB_RUN_EA_FACE")
CSC = os.environ.get("PB_CSC", "doNotCompress")  # characterSpacingControl
THEME_LANG = os.environ.get("PB_THEME_LANG")  # <w:themeFontLang w:eastAsia=...> in settings
FILL = os.environ.get("PB_FILL", "あ")  # the filler character  # rFonts eastAsia on the run itself  # "" = docDefaults names no eastAsia face
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
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="%s"/>%s'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>') % (W, CSC, '<w:themeFontLang w:val="en-US" w:eastAsia="%s"/>' % THEME_LANG if THEME_LANG else '', COMPAT)
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="%s" w:hAnsi="%s"%s/><w:sz w:val="20"/><w:lang w:val="en-US"%s/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/><w:jc w:val="%s"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') % (W, ASCII_FONT, ASCII_FONT, ' w:eastAsia="%s"' % DD_EA_FACE if DD_EA_FACE else "",
                ' w:eastAsia="%s"' % EA_LANG if EA_LANG else "", JC)
OPENING = ["｢", "「", "（", "(", "＄", "￥", "‘", "“"]
CLOSING = ["｣", "､", "｡", "･", "ﾞ", "ﾟ", "」", "、", "。", "）", "・", "ー", "っ", "々", "％", "℃", "’", "”", "!", "?", "%", ",", "."]
CELLS = 40  # 10pt cells in 400pt
KOFF = os.environ.get("PB_KINSOKU_OFF") == "1"  # <w:kinsoku w:val="0"/> on every paragraph


def para(text):
    return '<w:p><w:pPr>%s<w:ind w:right="%d"/></w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p><w:p/>' % ('<w:kinsoku w:val="0"/>' if KOFF else '', 9026 - CELLS * 200, '<w:rPr>%s<w:lang w:eastAsia="%s"/></w:rPr>' % ('<w:rFonts w:eastAsia="%s"/>' % RUN_EA_FACE if RUN_EA_FACE else '', RUN_LANG) if RUN_LANG else '', text)


def build():
    body = ""
    for m in OPENING:
        # the opening mark lands in the last cell (half-width marks take half a cell)
        body += para(FILL * (CELLS - 1) + m + "いうえお")
    for m in CLOSING:
        body += para(FILL * CELLS + m + "いうえお")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/></w:sectPr></w:body></w:document>'
           % (W, body))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "hk%s_c%d_%s_%s_%s.docx" % ("_off" if KOFF else "", COMPAT, JC, EA_LANG, ASCII_FONT.replace(" ", "")) if not RUN_LANG else "hk_c%d_%s_dd%s_%s_run%s%s%s.docx" % (COMPAT, JC, EA_LANG or "none", (DD_EA_FACE or "noface").replace(" ", ""), RUN_LANG + ("_f" if RUN_EA_FACE else ""), "_" + CSC if CSC != "doNotCompress" else "", "_" + FILL if FILL != "あ" else "", ) if not THEME_LANG else "hk_c%d_%s_run%s_theme%s.docx" % (COMPAT, JC, RUN_LANG, THEME_LANG))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    return path


def pairs(lines):
    out = []
    i = 0
    lines = [l.replace(" ", "") for l in lines if l.replace(" ", "")]
    while i + 1 < len(lines):
        if lines[i].startswith(FILL):
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
    p = build(); pdf = p[:-5] + ".pdf"
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    wp = pairs(word_lines(pdf))
    op = pairs(oxi_lines(os.environ["PB_OXI"], p)) if os.environ.get("PB_OXI") else [None] * len(wp)
    for m, a, b in zip(OPENING + CLOSING, wp, op):
        kind = "end" if m in OPENING else "start"
        print("%s %-5s Word %-14s Oxi %-14s %s" % (m, kind, a, b, "" if a == b else "<<"), flush=True)
