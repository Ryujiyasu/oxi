# -*- coding: utf-8 -*-
"""How much does a JUSTIFIED compat-15 line squeeze to keep its last unit -- on a MID line
of the paragraph, not only on the last one?

S1585 (`_pb_cjk_jshrink_gen.py`) swept one-paragraph texts, so every measured line was the
paragraph's LAST line (left-set even under jc=both). `_pb_kinsoku_off_squeeze_gen.py`
PB_HALF=6 then found a mid line keeping «ab» by its whole autospace gap (2.55pt at 10pt)
where S1585's gap term gives half of it (1.3), and a paragraph-final «ab» wrapping instead.

Arms: the unit that overflows (Latin word with its own leading gap / CJK character with
k Latin islands earlier on the line / no gap at all) x the measured line = paragraph's last
line or a mid line (a CJK tail follows) x the ascii face (Century, or ＭＳ 明朝 with
PB_LATIN=ＭＳ 明朝). One document per (arm, jc): 400 paragraphs whose right indent grows
one twip at a time; readout by COM, the line number of the unit's first character vs the
paragraph's first character (wdFirstCharacterLineNumber). allow = left flip - both flip.
Each document opens in its own Word process.

    python tools/metrics/_pb_jshrink_mid_gen.py [arm ...]
"""
import os, re, sys, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = Path("pipeline_data/_pb_jshrink_mid"); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
LATIN = os.environ.get("PB_LATIN", "Century")
BALANCE = os.environ.get("PB_BALANCE", "0") == "1"
SZ = int(float(os.environ.get("PB_PT", "10.5")) * 2)


def settings():
    b = '<w:balanceSingleByteDoubleByteWidth/>' if BALANCE else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat>%s<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>' % (W, b))


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="%s" w:eastAsia="ＭＳ 明朝" w:hAnsi="%s" w:cs="Times New Roman"/>'
          '<w:kern w:val="2"/><w:sz w:val="%d"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>') % (W, LATIN, LATIN, SZ)
SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="365"/></w:sectPr>')
TEXT_W = 9638
N = int(os.environ.get("NPARA", "400"))
F = "新潟県新潟市中央区西堀前通新潟市中央区西堀前通新潟県新潟市中央区西堀前通新潟"  # 38 cells
TAIL = "新潟県新潟市中央区西堀前通新潟市中央区"
# (text, index of the unit's first character)
ARMS = {
    "lat_g1":  (F + "市AB", 39),                                   # Latin word, its own gap before it
    "lat_g0":  (F + "、AB", 39),                                   # Latin word after 、 (no gap)
    "cjk_k1":  (F[:19] + "AB" + F[19:] + "市", 40),               # CJK unit, one island (2 gaps) mid-line
    "cjk_k2":  (F[:12] + "AB" + F[12:26] + "AB" + F[26:], 40),    # CJK unit, two islands (4 gaps)
    "cjk_g0":  (F + "市中", 39),                                   # pure CJK (control)
    "lat_g3":  (F[:19] + "AB" + F[19:] + "市AB", 42),             # Latin word with its own gap + an island earlier (3 gaps)
    "lat_long": (F[:36] + "ABCDEF", 36),                          # longer Latin word with its own gap (the 0.35 x word ceiling?)
}


def runs(text):
    out = []
    for m in re.finditer(r"[0-9A-Za-z-]+|[^0-9A-Za-z-]+", text):
        t = m.group(0)
        hint = '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>' if (t[0].isascii() and t[0].isalnum()) else ""
        out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (hint, t))
    return "".join(out)


def build(arm, mid, jc, w0_tw):
    text = ARMS[arm][0] + (TAIL if mid else "")
    ps = []
    for i in range(N):
        right = TEXT_W - (w0_tw - i)
        ps.append('<w:p><w:pPr><w:ind w:right="%d"/><w:jc w:val="%s"/></w:pPr>%s</w:p>' % (right, jc, runs(text)))
    doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s%s</w:body></w:document>' % (W, "".join(ps), SECT)
    at = OUT / ("%s_%s_%s_%s_bal%d.docx" % (arm, "mid" if mid else "last", jc, LATIN.replace(" ", ""), int(BALANCE)))
    with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RR); z.writestr("word/_rels/document.xml.rels", DR)
        z.writestr("word/settings.xml", settings()); z.writestr("word/styles.xml", STYLES); z.writestr("word/document.xml", doc)
    return at


def flip(arm, mid, jc, w0_tw):
    """The last width (pt) at which the unit still sits on the paragraph's first line."""
    import win32com.client
    at = build(arm, mid, jc, w0_tw)
    pos = ARMS[arm][1]
    app = win32com.client.DispatchEx("Word.Application"); app.Visible = False; app.DisplayAlerts = 0
    try:
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            for i in range(N):
                p = d.Paragraphs(i + 1).Range
                first = p.Characters(1).Information(10)
                unit = p.Characters(pos + 1).Information(10)
                if unit != first:
                    return (w0_tw - i + 1) / 20.0
            return None
        finally:
            d.Close(False)
    finally:
        app.Quit()


if __name__ == "__main__":
    arms = sys.argv[1:] or list(ARMS)
    for arm in arms:
        for mid in (False, True):
            # start the left sweep 10pt above the text's rough natural width (CJK cells + half
            # cells for Latin + a quarter em per CJK|Latin boundary), so the window holds the flip
            first_line = ARMS[arm][0]
            cells = sum(0.5 if c.isascii() else 1.0 for c in first_line)
            gaps = sum(1 for a, b in zip(first_line, first_line[1:]) if a.isascii() != b.isascii())
            w0 = int(round((cells * SZ / 2.0 + gaps * SZ / 8.0 + 10.0) * 20))
            left = flip(arm, mid, "left", min(w0, TEXT_W))
            if left is None:
                print(arm, "mid" if mid else "last", "left never wraps in window"); continue
            both = flip(arm, mid, "both", int(round(left * 20)) + 20)
            print("%-8s %-4s %s bal%d natural(left)=%7.2f justified min=%s allow=%s" % (
                arm, "mid" if mid else "last", LATIN, int(BALANCE), left,
                None if both is None else round(both, 2), None if both is None else round(left - both, 2)), flush=True)
