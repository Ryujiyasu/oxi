# -*- coding: utf-8 -*-
"""Does Word bill a proportional face's brackets wider at BREAK time than it draws them?

ja/educational__08709ff2 «②対象の地域（ここでは東京都江東区）を押し「バツ印」…» (BIZ
UDPGothic 18pt): on the faithful slice Word pushes «印」» down while the drawn line would
still end 1.28pt inside the edge; the «…調べる方» line of the same document (no
brackets) keeps 方 with 0.2pt to spare.  Hypothesis: each bracket costs more at break
time than its 9.0pt advance.

Arms: k pairs «（あ）» then kana, then a final kanji 亜 followed by more text; right
indent swept 0.2pt at a time.  Per arm: the smallest drawn slack (edge - end of 亜)
at which Word still keeps 亜 on line 1, and the largest at which it pushes it.

    python tools/metrics/_pb_bracket_fit_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_bracket_fit")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
FACE = os.environ.get("PB_FACE", "BIZ UDPGothic")
KANA = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめも"
BR = os.environ.get("PB_BR", "（あ）")
CM = int(os.environ.get("PB_CM", "11"))
RIGHTS = list(range(0, 480, 4))
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')


def _adv():
    from fontTools.ttLib import TTCollection
    import glob
    for f in glob.glob(r"C:/Windows/Fonts/*.ttc"):
        for t in TTCollection(f).fonts:
            if FACE in {n.toUnicode() for n in t["name"].names if n.nameID in (1, 4)}:
                cm = t.getBestCmap()
                return lambda c: t["hmtx"][cm[ord(c)]][0] * 18 / t["head"].unitsPerEm
    raise SystemExit("face not found")


def text(k):
    adv = _adv()
    t = BR * k
    i = 0
    # the text before 亜 ends ~10pt inside the full line (425.2), so the sweep (24pt) crosses it
    while sum(map(adv, t)) + adv(KANA[i % len(KANA)]) < float(os.environ.get("PB_FILL", "407")):
        t += KANA[i % len(KANA)]; i += 1
    return t + "亜" + "まみむめも"


def build(k):
    ps = []
    for n, r in enumerate(RIGHTS):
        ps.append('<w:p><w:pPr>%s<w:ind w:right="%d"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
                  '<w:sz w:val="36"/><w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>%s</w:t></w:r></w:p>'
                  % ("<w:pageBreakBefore/>" if n else "", r, FACE, FACE, FACE, text(k)))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1701" w:right="1701"/></w:sectPr>'
           '</w:body></w:document>' % (W, "".join(ps)))
    st = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="compressPunctuation"/><w:compat>'
          '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>' % (W, CM))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "bf_%s_k%d_%x_c%d.docx" % (FACE.replace(" ", ""), k, ord(BR[0]), CM))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/settings.xml", st)
        z.writestr("word/document.xml", doc)
    return path


def verdict(pdf):
    import fitz
    keep, push = [], []
    for r, pg in zip(RIGHTS, fitz.open(pdf)):
        edge = (11906 - 1701 - r) / 20
        lines = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        if c["c"].strip():
                            lines.setdefault(round(c["origin"][1]), []).append(c)
        first = sorted(lines[min(lines)], key=lambda c: c["origin"][0])
        second = sorted(lines[sorted(lines)[1]], key=lambda c: c["origin"][0]) if len(lines) > 1 else []
        if first[-1]["c"] == "亜":
            keep.append(round(edge - first[-1]["bbox"][2], 2))
        elif second and second[0]["c"] == "亜":
            # drawn slack had 亜 stayed: its advance after the last kept glyph
            push.append(round(edge - (first[-1]["bbox"][2] + 18.0), 2))
    return keep, push


if __name__ == "__main__":
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for k in map(int, os.environ.get("PB_KS", "0,1,2,4").split(",")):
            docx = build(k); pdf = docx[:-5] + ".pdf"
            d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
            keep, push = verdict(pdf)
            print("%s c%d k=%d %r  keep min slack %s   push max slack %s" % (FACE, CM, k, BR, min(keep) if keep else None, max(push) if push else None), flush=True)
    finally:
        w.Quit()
