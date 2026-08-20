# -*- coding: utf-8 -*-
"""What do a Symbol bullet and HTML autospacing add to an image line?

★PREMISE CORRECTED (2026-08-20, second session): "every figure host is a
Symbol-bullet LIST item (numId 7) with beforeAutospacing" is FALSE — the real
creative__0158c02a has that form ONCE (p1058 "Amir Siraj"); the other 28 image
hosts are PLAIN paragraphs (pPr = w:line only, docDefaults after=160).  Their
−1.4/block was the S1180 double-subtraction (see _pb_crfig_gen.py).  The
autospacing×image amounts below are still real Word derivations (numaut_img
6.75 / aut_img 14.625 vs Oxi's flat 5.625, the S931 before-only asymmetry, the
same-list blk collapse) — but their corpus exposure is that single block, so
they are bug-mine material, not the creative lever.

Original framing: this probe decomposes marker growth on an IMAGE line
([[word_marker_line_height_law]]'s image-line variant), the body autospacing
amount @276, and their combination, plus one arm shaped like the
[caption][sp+img] block.

    python _pb_imgnum_gen.py gen
    python _pb_imgnum_gen.py read     # Word COM truth
    python _pb_imgnum_gen.py oxi [ENV=..]
"""
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_imgnum")
DOCX = os.path.join(OUT, "imgnum.docx")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_brimg_gen import PNG, pic  # noqa: E402  (1x1 png + wp:inline builder)

FONT, SZ = "Calibri", 22            # 11pt, natural 13.4277
REPEAT = 8

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/img.png"/>'
         '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
         '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/>'
          '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '</w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/></w:style></w:styles>')
# the specimen's lvl-0: Symbol bullet U+F0B7, ind left=786 hanging=360.
# numId 6 = a SECOND list with the same bullet (creative's captions are numId 6
# while its figures are numId 7 — the cross-list adjacency is part of the spec).
NUMBERING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + NS + '>'
             '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
             '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
             '<w:lvlText w:val=""/><w:lvlJc w:val="left"/>'
             '<w:pPr><w:ind w:left="786" w:hanging="360"/></w:pPr>'
             '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>'
             '</w:lvl></w:abstractNum>'
             '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>'
             '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
             '<w:lvlText w:val=""/><w:lvlJc w:val="left"/>'
             '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
             '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>'
             '</w:lvl></w:abstractNum>'
             '<w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num>'
             '<w:num w:numId="6"><w:abstractNumId w:val="1"/></w:num></w:numbering>')

# (label, kind, numpr, auto_b, auto_a, img_h, opts)
#   kind: "txt" = one text line; "img" = [space + inline img]; "blk" =
#   [caption txt para][sp+img para] per copy (the real creative block).
#   opts: sz (half-points), line, cap_num (blk caption's numId; fig is 7),
#   both default to the module constants.
ARMS = [
    ("txt@276",        "txt", False, False, False, 0,    {}),
    ("aut_txt@276",    "txt", False, True,  True,  0,    {}),
    ("num_txt@276",    "txt", True,  False, False, 0,    {}),
    ("numaut_txt@276", "txt", True,  True,  True,  0,    {}),
    ("img36@276",      "img", False, False, False, 36.0, {}),
    ("num_img36@276",  "img", True,  False, False, 36.0, {}),
    ("aut_img36@276",  "img", False, True,  False, 36.0, {}),
    ("numaut_img36",   "img", True,  True,  False, 36.0, {}),
    ("numaut_img60",   "img", True,  True,  False, 60.0, {}),
    ("blk36@276",      "blk", True,  True,  False, 36.0, {}),
    # ---- second cut (amounts-model sweep) ----
    ("autB_txt@276",   "txt", False, True,  False, 0,    {}),          # before only
    ("autA_txt@276",   "txt", False, False, True,  0,    {}),          # after only
    ("autB_txt@240",   "txt", False, True,  False, 0,    {"line": 240}),
    ("autB_txt14",     "txt", False, True,  False, 0,    {"sz": 28}),  # 14pt
    ("numautB_txt",    "txt", True,  True,  False, 0,    {}),          # list, before only
    ("numaut_img14",   "img", True,  True,  False, 36.0, {"sz": 28}),
    ("numaut_img@240", "img", True,  True,  False, 36.0, {"line": 240}),
    ("aut_img@240",    "img", False, True,  False, 36.0, {"line": 240}),
    ("blk36_x67",      "blk", True,  True,  False, 36.0, {"cap_num": 6}),  # creative: cap 6 / fig 7
    ("blk36_nolist",   "blk", False, True,  False, 36.0, {}),
]


def rpr(font=FONT, sz=SZ):
    return ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (font, font, font, sz, sz))


def ppr(pbb=False, numpr=False, auto_b=False, auto_a=False, line=276, sz=SZ, numid=7):
    num = ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>' % numid
           if numpr else "")
    sp = ('<w:spacing w:before="100" w:beforeAutospacing="%d" w:after="100"'
          ' w:afterAutospacing="%d" w:line="%d" w:lineRule="auto"/>'
          % (1 if auto_b else 0, 1 if auto_a else 0, line)
          if (auto_b or auto_a) else
          '<w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="auto"/>' % line)
    return ("<w:pPr>%s<w:widowControl w:val=\"0\"/>%s%s%s</w:pPr>"
            % ("<w:pageBreakBefore/>" if pbb else "", num, sp, rpr(sz=sz)))


def subject(kind, numpr, auto_b, auto_a, h, idx, opts):
    sz = opts.get("sz", SZ)
    line = opts.get("line", 276)
    if kind == "txt":
        return ('<w:p>%s<w:r>%s<w:t xml:space="preserve">body text line</w:t></w:r></w:p>'
                % (ppr(numpr=numpr, auto_b=auto_b, auto_a=auto_a, line=line, sz=sz),
                   rpr(sz=sz)))
    img_p = ('<w:p>%s<w:r>%s<w:t xml:space="preserve"> </w:t></w:r>'
             "<w:r>%s%s</w:r></w:p>"
             % (ppr(numpr=numpr, auto_b=auto_b, auto_a=False, line=line, sz=sz),
                rpr(sz=sz), rpr(sz=sz), pic(idx, h)))
    if kind == "img":
        return img_p
    cap_p = ('<w:p>%s<w:r>%s<w:t xml:space="preserve">Figure caption text</w:t></w:r></w:p>'
             % (ppr(numpr=numpr, auto_b=True, auto_a=True, line=line, sz=sz,
                    numid=opts.get("cap_num", 7)), rpr(sz=sz)))
    return cap_p + img_p


def marker(tag, pbb=False):
    return ('<w:p>%s<w:r>%s<w:t>%s</w:t></w:r></w:p>'
            % (ppr(pbb=pbb, line=240), rpr(), tag))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body, idx = [], 100
    body.append(marker("M00S", pbb=True))
    body.append(marker("M00E"))
    for ai, (_lbl, kind, numpr, ab, aa, h, opts) in enumerate(ARMS, start=1):
        body.append(marker("M%02dS" % ai, pbb=True))
        for _ in range(REPEAT):
            idx += 1
            body.append(subject(kind, numpr, ab, aa, h, idx, opts))
        body.append(marker("M%02dE" % ai))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>')
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", NUMBERING)
        z.writestr("word/media/img.png", PNG)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(ARMS), "arms x", REPEAT)


def report(spans, who):
    base = spans.get(0)
    if base is None:
        raise SystemExit("control arm missing")
    print("%s  control span (one marker) = %.2f" % (who, base))
    print("%-16s %9s" % ("arm", "per copy"))
    for ai, (lbl, *_rest) in enumerate(ARMS, start=1):
        span = spans.get(ai)
        print("%-16s %9s" % (lbl, "MISSING" if span is None
                             else "%.3f" % ((span - base) / REPEAT)))


def read():
    import re
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.ScreenUpdating = False
    d = app.Documents.Open(DOCX, ReadOnly=True)
    ys = {}
    try:
        d.Repaginate()
        for p in d.Paragraphs:
            rng = p.Range
            m = re.match(r"M(\d\d)([SE])", rng.Text)
            if not m:
                continue
            c = d.Range(rng.Start, rng.Start)
            ys[(int(m.group(1)), m.group(2))] = (c.Information(3), round(c.Information(6), 2))
    finally:
        d.Close(False)
        app.Quit()
    spans = {}
    for ai in range(0, len(ARMS) + 1):
        s, e = ys.get((ai, "S")), ys.get((ai, "E"))
        if s and e and s[0] == e[0]:
            spans[ai] = e[1] - s[1]
    report(spans, "WORD")


def oxi(envs=""):
    import json
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "imgnum_oxi.json")
    subprocess.run([GDI, DOCX, os.path.join(tempfile.gettempdir(), "imgnum"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    ys = {}
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        for e in pg["elements"]:
            t = (e.get("text") or "").strip()
            if len(t) == 4 and t.startswith("M") and t[3] in "SE" and t[1:3].isdigit():
                ys.setdefault((int(t[1:3]), t[3]), (pg["page"], e["y"]))
    spans = {}
    for ai in range(0, len(ARMS) + 1):
        s, e = ys.get((ai, "S")), ys.get((ai, "E"))
        if s and e and s[0] == e[0]:
            spans[ai] = e[1] - s[1]
    report(spans, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "read": read}[sys.argv[1]]()
