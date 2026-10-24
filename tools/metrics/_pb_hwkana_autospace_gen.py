# -*- coding: utf-8 -*-
"""Does Word put the autoSpaceDE gap between HALF-WIDTH katakana and Latin letters?

`_pb_kinsoku_off_squeeze_gen.py` PB_HALF=4: «あ×37 + 、ｱ + abc» fills a 40-cell line to the
twip in Oxi (ｱ half a cell, no gap) but Word wraps «abc» at compat 14 -- as if a quarter-em
gap sat between ｱ and a, which Oxi's is_cjk_latin_joint (kana ranges only) never inserts.

Arms: the character before the Latin word = half-width ｱ / full-width カ / ASCII x (no gap
control), autoSpaceDE on / off, compat 14 / 15.  Each paragraph «あ×36 + <c> + abc + いうえお»
sits short of the line end so the gap can be read directly from the PDF glyph origins:
gap = x(a) - x(c) - advance(c).  Prints the gap in em for Word and Oxi.

    python tools/metrics/_pb_hwkana_autospace_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_hwkana_autospace")
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
FONT = os.environ.get("PB_FONT", "ＭＳ 明朝")
CHARS = {"hwkana": "ｱ", "fwkana": "カ", "ascii": "x"}
SIZE = 10.0


def build(compat, autospace):
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="compressPunctuation"/>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>') % (W, compat)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
              '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/><w:jc w:val="left"/>%s</w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
              ) % (W, FONT, FONT, FONT, int(SIZE * 2), "" if autospace else '<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>')
    body = "".join('<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p><w:p/>' % ("あ" * 10 + c + "abc" + "いうえお") for c in CHARS.values())
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/></w:sectPr></w:body></w:document>'
           % (W, body))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "hk_c%d_%s_%s.docx" % (compat, "as" if autospace else "noas", FONT.replace(" ", "")))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    return path


def gaps_word(pdf):
    import pymupdf
    out = {}
    pg = pymupdf.open(pdf)[0]
    rows = {}
    for b in pg.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    rows.setdefault(round(c["origin"][1]), []).append((c["origin"][0], c["c"], c["bbox"][2]))
    for k in sorted(rows):
        r = sorted(rows[k])
        t = "".join(c for _, c, _ in r)
        for name, ch in CHARS.items():
            i = t.find(ch + "abc")
            if i >= 0:
                xc, xa = r[i][0], r[i + 1][0]
                # the character's advance = distance to the next glyph when no gap: use あ pitch for
                # full-width, half of it for half-width / ASCII
                pitch = r[1][0] - r[0][0]
                adv = pitch if name == "fwkana" else pitch / 2
                out[name] = (xa - xc - adv) / SIZE
    return out


def gaps_oxi(exe, docx):
    import subprocess, json, tempfile
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    out = {}
    for p in json.load(open(g, encoding="utf-8"))["pages"]:
        rows = {}
        for c in p["glyphs"]:
            rows.setdefault(round(c["baseline"]), []).append(c)
        for k in sorted(rows):
            r = sorted(rows[k], key=lambda c: c["x"])
            t = "".join(c["char"] for c in r)
            for name, ch in CHARS.items():
                i = t.find(ch + "abc")
                if i >= 0:
                    pitch = r[1]["x"] - r[0]["x"]
                    adv = pitch if name == "fwkana" else pitch / 2
                    out[name] = (r[i + 1]["x"] - r[i]["x"] - adv) / SIZE
    return out


if __name__ == "__main__":
    import win32com.client
    for compat in (15, 14):
        for autospace in (True, False):
            p = build(compat, autospace); pdf = p[:-5] + ".pdf"
            w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
            try:
                d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
            finally:
                w.Quit()
            gw = gaps_word(pdf)
            go = gaps_oxi(os.environ["PB_OXI"], p) if os.environ.get("PB_OXI") else {}
            for name in CHARS:
                print("c%d autospace=%-5s %-7s gap/em Word %6.3f  Oxi %s" % (
                    compat, autospace, name, gw.get(name, float("nan")),
                    ("%6.3f" % go[name]) if name in go else "  -"), flush=True)
