# -*- coding: utf-8 -*-
"""Where does Word break «「·MONO」» when the line ends inside it?

ja/technical__c6d7cdb1 p18 (compat 14, run lang eastAsia en-US): Word ends a line with
«…または「·» and starts the next with «MONO」設定…» -- a break between the middle dot
(U+00B7) and the Latin letters.  Oxi keeps «·MONO» as one word and wraps it whole.

Arms: compat 14 / 15 x run eastAsia lang ja-JP / en-US x the middle dot U+00B7 / the
katakana middle dot U+30FB / a hyphen-minus.  Each paragraph: N «あ» + «「·MONO」いう»
with N stepped so the line end falls at every position inside «「·MONO」».  Prints, per
arm and N, the first line's tail and the second line's head, Word and Oxi.

    python tools/metrics/_pb_middot_latin_break_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_middot_latin_break")
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
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="20"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/><w:jc w:val="left"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') % W
CELLS = 40
DOTS = {"middot": "·", "katakana": "・", "hyphen": "-"}
HINT = os.environ.get("PB_HINT") == "1"
NOOPEN = os.environ.get("PB_NOOPEN") == "1"  # drop the «「»: does the hinted dot itself start a line?  # <w:rFonts w:hint="eastAsia"/> on the run (c6d7cdb1 carries it)


def build(compat, lang, dot_name):
    dot = DOTS[dot_name]
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="compressPunctuation"/>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>') % (W, compat)
    body = ""
    # «「·MONO」» spans 1 + 0.5 + 4 x ~0.5 + 1 cells; sweep N so the line end lands at each joint
    for n in range(CELLS - 5, CELLS):
        body += ('<w:p><w:pPr><w:ind w:right="%d"/></w:pPr><w:r><w:rPr><w:lang w:eastAsia="%s"/></w:rPr>'
                 '<w:t xml:space="preserve">%s</w:t></w:r></w:p><w:p/>'
                 % (9026 - CELLS * 200, lang, "あ" * (n + 1 if NOOPEN else n) + ("" if NOOPEN else "「") + dot + "MONO」いうえお")).replace(
                     '<w:rPr>', '<w:rPr><w:rFonts w:hint="eastAsia"/>' if HINT else '<w:rPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/></w:sectPr></w:body></w:document>'
           % (W, body))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "md_c%d_%s_%s%s.docx" % (compat, lang, dot_name, "_hint" if HINT else "", ) if not NOOPEN else "md_c%d_%s_%s%s_noopen.docx" % (compat, lang, dot_name, "_hint" if HINT else ""))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    return path


def pairs(lines):
    out, i = [], 0
    lines = [l.replace(" ", "") for l in lines if l.replace(" ", "")]
    while i + 1 < len(lines):
        if lines[i].startswith("あ"):
            out.append((lines[i].lstrip("あ"), lines[i + 1][:4])); i += 2
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
    for compat in (14, 15):
        for lang in ("en-US", "ja-JP"):
            for dot_name in os.environ.get("PB_DOTS", "middot,katakana,hyphen").split(","):
                p = build(compat, lang, dot_name); pdf = p[:-5] + ".pdf"
                w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
                try:
                    d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
                finally:
                    w.Quit()
                wp = pairs(word_lines(pdf))
                op = pairs(oxi_lines(os.environ["PB_OXI"], p)) if os.environ.get("PB_OXI") else [None] * len(wp)
                for n, a, b in zip(range(CELLS - 5, CELLS), wp, op):
                    print("c%d %-5s %-8s n=%d Word %-18s Oxi %-18s %s" % (compat, lang, dot_name, n, a, b, "" if a == b else "<<"), flush=True)
