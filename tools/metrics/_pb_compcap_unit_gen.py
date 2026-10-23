# -*- coding: utf-8 -*-
"""Re-measure S1490's compat-15 punctuation capacity on this machine's Word, including the
unit X + 、 (an overflowing character followed by a line-start-prohibited mark).

S1490's table (old machine, compcap.py, ＭＳ 明朝 10.5, jc both, docGrid lines):
normal X n=40 k=1..4: 0.452 0.624 0.714 0.762 cell; unit X+mark 0.429 0.538 0.614.
The new machine's truth refuses units at 0.56-0.59 with k=3-4 (ja/administrative__0ba21b0c,
ja/legal__08b9b46a, ja/policies__0d079ffa) that the table packs.

Each paragraph: k marks spread over the line, kana up to N-1 characters, then X «字»
(arm unit: followed by «、» + more text). The right indent sweeps in 6tw steps; prints the
smallest overflow (cells) at which X leaves line 1.

    python tools/metrics/_pb_compcap_unit_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_compcap_unit")
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
FACE = os.environ.get("PB_FACE", "ＭＳ 明朝")
SZ = int(os.environ.get("PB_SZ", "21"))
CELL = SZ * 10  # twips
TEXTW = 11906 - 2 * 1440
# PB_LINE2=1: the measured line is the paragraph's SECOND line (a short first line + w:br)
LINE2 = os.environ.get("PB_LINE2") == "1"
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:balanceSingleByteDoubleByteWidth/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % W
MARK = os.environ.get("PB_MARK", "、")
KS = [int(v) for v in os.environ.get("PB_K", "1,2,3,4").split(",")]
KANA = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわ"


def line(n, k, unit):
    chars = []
    step = max(1, (n - 2) // (k + 1)) if k else n
    ki = 0
    for i in range(n - 1):
        if k and ki < k and i > 0 and i % step == 0:
            chars.append(MARK); ki += 1
        else:
            chars.append(KANA[i % len(KANA)])
    return "".join(chars) + "字" + ("、つづきの文です。" if unit else "つづきの文です。")


def build(n):
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
              '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') % (W, FACE, FACE, FACE, SZ)
    body = ""
    for unit in (False, True):
        for k in KS:
            txt = line(n, k, unit)
            for over in range(0, CELL, int(os.environ.get("PB_STEP", "6"))):
                ind = TEXTW - (n * CELL - over)
                tag = "U%d%d_%d" % (int(unit), k, over)
                body += ('<w:p><w:pPr><w:ind w:right="%d"/></w:pPr>%s<w:r><w:t>%s</w:t></w:r></w:p><w:p><w:r><w:t>%s</w:t></w:r></w:p>'
                         % (ind, '<w:r><w:t>前の行</w:t><w:br/></w:r>' if LINE2 else '', txt, tag))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/>'
           '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>' % (W, body))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "cc_n%d_s%d%s%s.docx" % (n, SZ, "_l2" if LINE2 else "", "_p" if MARK != "、" else ""))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    return path


def flips(lines):
    """tag -> X stayed on line 1"""
    out, cur = {}, []
    for t in lines:
        t = t.strip()
        if t.startswith("U") and "_" in t and t[1:3].isdigit():
            first = [c for c in cur if c not in ("前の行", "前", "の", "行")]
            out[t] = bool(first) and "字" in first[0]
            cur = []
        elif t:
            cur.append(t)
    return out


def word_lines(pdf):
    import pymupdf
    res = []
    for pg in pymupdf.open(pdf):
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                res.append("".join(s["text"] for s in l["spans"]))
    return res


def oxi_lines(exe, docx):
    import subprocess, json, tempfile
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    res = []
    for p in json.load(open(g, encoding="utf-8"))["pages"]:
        rows = {}
        for c in p["glyphs"]:
            rows.setdefault(round(c["baseline"], 1), []).append(c)
        for k in sorted(rows):
            res.append("".join(c["char"] for c in sorted(rows[k], key=lambda c: c["x"])))
    return res


def report(name, fl):
    for unit in (0, 1):
        row = []
        for k in KS:
            overs = sorted(int(t.split("_")[1]) for t in fl if t.startswith("U%d%d_" % (unit, k)))
            first_refuse = next((o for o in overs if not fl["U%d%d_%d" % (unit, k, o)]), None)
            row.append("k=%d %s" % (k, "%.3f" % (first_refuse / CELL) if first_refuse is not None else "-"))
        print("%-5s %s  %s" % (name, "unit" if unit else "X   ", "  ".join(row)), flush=True)


if __name__ == "__main__":
    import win32com.client
    exe = os.environ.get("PB_OXI")
    for n in [int(v) for v in os.environ.get("PB_N", "40,26").split(",")]:
        p = build(n); pdf = p[:-5] + ".pdf"
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
        print("n=%d sz=%d first overflow (cells) at which X wraps:" % (n, SZ))
        report("Word", flips(word_lines(pdf)))
        if exe:
            report("Oxi", flips(oxi_lines(exe, p)))
