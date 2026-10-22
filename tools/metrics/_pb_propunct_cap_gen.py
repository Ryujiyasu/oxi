# -*- coding: utf-8 -*-
"""How much room do proportional punctuation marks lend a full line (compressPunctuation)?

ja/educational__0d3e20fe (Meiryo UI 20pt, compat 14, compressPunctuation, jc both): Word
sets «ーです。文字の大きさ、量、字間、行間等を確認する» on one line -- 15.98pt over the
415.3pt text width -- and squeezes the whole line. S1207 (derived on MS PMincho, whose
«、» is already ~0.5em) gives proportional faces NO punctuation credit; Meiryo UI's «、»
is 0.7em. Hypothesis: the credit is Σ (mark advance - 0.5em) for every face.

One document per arm (fresh Word process): the same paragraph repeated with the right
indent stepped by 5tw; each paragraph = n marks + filler ideographs + one probe kana at the
end. Reads, per indent, whether the last kana stayed on line 1; prints the flip indent and
the implied credit (natural width - text width at the flip).

    python tools/metrics/_pb_propunct_cap_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_propunct_cap")
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
FACES = {"meiryoui": "Meiryo UI", "pmincho": "ＭＳ Ｐ明朝", "mincho": "ＭＳ 明朝"}
COMPAT = os.environ.get("PB_COMPAT", "14")
JC = os.environ.get("PB_JC", "both")
SZ = int(os.environ.get("PB_SZ", "20"))  # half-points
GRID = os.environ.get("PB_GRID")  # linePitch for a lines docGrid
FILL = "文字間行等確認大量章"


def para(n, ind):
    marks = "、" * 0
    body = ""
    k = 0
    for i in range(n):
        body += FILL[k % len(FILL)] + "、"; k += 1
    while len(body) < int(os.environ.get("PB_LEN", "41")):
        body += FILL[k % len(FILL)]; k += 1
    body += "る次行"
    return ('<w:p><w:pPr><w:ind w:right="%d"/><w:jc w:val="%s"/></w:pPr><w:r><w:t>%s</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>IND%d</w:t></w:r></w:p>' % (ind, JC, body, ind))


def build(face, n):
    f = FACES[face]
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
              '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') % (W, f, f, f, SZ)
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="compressPunctuation"/>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%s"/></w:compat></w:settings>') % (W, COMPAT)
    body = "".join(para(n, ind) for ind in range(0, 1600, 5))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/>%s</w:sectPr></w:body></w:document>' % (W, body, '<w:docGrid w:type="lines" w:linePitch="%s"/>' % GRID if GRID else ""))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "pc_%s_n%d_c%s_%s_s%d%s.docx" % (face, n, COMPAT, JC, SZ, "_g" + GRID if GRID else ""))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    return path


def read(pdf):
    """indent -> (line-1 text, natural advances of line 1 chars measured where nothing squeezes)"""
    import pymupdf
    lines = []
    for pg in pymupdf.open(pdf):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                ch = [c for s in l["spans"] for c in s["chars"]]
                lines.append(("".join(c["c"] for c in ch), ch))
    out = {}
    cur = []
    for t, ch in lines:
        if t.startswith("IND"):
            out[int(t[3:])] = cur; cur = []
        else:
            cur.append((t, ch))
    return out


if __name__ == "__main__":
    import win32com.client
    arms = [(f, n) for f in os.environ.get("PB_FACES", ",".join(FACES)).split(",") for n in (0, 2, 4)]
    for face, n in arms:
        p = build(face, n); pdf = p[:-5] + ".pdf"
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
        res = read(pdf)
        # natural line: widest indent's... use the indent where line 1 = all incl. る first time (from large indent down)
        flip = None
        for ind in sorted(res):
            l1 = res[ind][0][0] if res[ind] else ""
            if "る" not in l1:
                flip = ind; break
        # natural advances from a squeeze-free line: the indent 0 line's chars
        ch = res[0][0][1]
        adv = {}
        for a, b in zip(ch, ch[1:]):
            adv.setdefault(a["c"], round(b["origin"][0] - a["origin"][0], 3))
        print("%-9s n=%d  first indent where «る» leaves line 1: %s tw   adv 、=%s 文=%s" % (face, n, flip, adv.get("、"), adv.get("文")), flush=True)
