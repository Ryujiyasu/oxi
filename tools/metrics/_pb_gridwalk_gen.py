# -*- coding: utf-8 -*-
"""Per-paragraph walk under a no-type docGrid: pin the cursor quantizer.

The blk-verbatim mini (2026-08-20, see ra_archive S910/S1180 notes) showed a
no-type `<w:docGrid w:linePitch="360"/>` makes Word quantize the CUMULATIVE
paragraph cursor to whole 96dpi pixels (0.75pt) at every paragraph boundary —
round on 6 boundaries, one image boundary that needed ceil.  Arm-style span
probes average the quantization away, so this probe reads EVERY paragraph's
collapsed-start Information(6).

Groups (each starts a fresh page via a pageBreakBefore marker, subjects have
after=0 via the Normal style, like creative__0158c02a):
  A: 10x 1-line text, Arial 12, line=276  (advance 15.868 = 21.157px — the
     cumulative fraction walks .16 .31 .47 .63 .79 .94 .10 ... a full cycle)
  B: 10x empty, same style               (empty path quantizer)
  C: [sp+img] line=360 with img_h 36 / 36.2 / 36.4 / 36.6, each followed by
     one text line (image-boundary fraction sweep: ceil vs round is +-0.75
     visible in the follower's y)
  D: 10x 1-line text, Arial 12, line=240 (f=1 — blk markers suggested the
     quantizer applies at single spacing too)
  E: 4x 1-line text line=276 with DIRECT after=160 (is the quantize applied
     to line+after together at the boundary?)

    python _pb_gridwalk_gen.py gen [pitch]   # default linePitch=360
    python _pb_gridwalk_gen.py read
    python _pb_gridwalk_gen.py oxi [ENV=..]
"""
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_gridwalk")
DOCX = os.path.join(OUT, "gridwalk.docx")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_brimg_gen import PNG, pic  # noqa: E402

FONT, SZ = "Arial", 24

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
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/img.png"/>'
         '</Relationships>')
# creative's shape: docDefaults after=160/line=259, Normal DECLARES after=0
# line=240 (paragraph style > docDefaults), so bare subjects resolve after=0.
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/>'
          '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>'
          '</w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/>'
          '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '</w:style></w:styles>')


def rpr():
    return ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (FONT, FONT, FONT, SZ, SZ))


def ppr(line, pbb=False, after=None):
    sp = ('<w:spacing %sw:line="%d" w:lineRule="auto"/>'
          % (('w:after="%d" ' % after) if after is not None else "", line))
    return ("<w:pPr>%s<w:widowControl w:val=\"0\"/>%s%s</w:pPr>"
            % ("<w:pageBreakBefore/>" if pbb else "", sp, rpr()))


def txt(line, text, after=None):
    return ('<w:p>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (ppr(line, after=after), rpr(), text))


def empty(line):
    return "<w:p>%s</w:p>" % ppr(line)


def img_p(line, h, idx):
    return ('<w:p>%s<w:r>%s<w:t xml:space="preserve"> </w:t></w:r><w:r>%s%s</w:r></w:p>'
            % (ppr(line), rpr(), rpr(), pic(idx, h)))


def marker(tag):
    return ('<w:p>%s<w:r>%s<w:t>%s</w:t></w:r></w:p>'
            % (ppr(240, pbb=True), rpr(), tag))


# manifest of (kind, detail) per paragraph, filled by gen(), used by read/oxi
def build():
    body, manifest, idx = [], [], 100
    def add(x, kind, detail=""):
        body.append(x)
        manifest.append((kind, detail))
    add(marker("GA"), "M", "GA")
    for i in range(10):
        add(txt(276, "walk text %d" % i), "txt276", str(i))
    add(marker("GB"), "M", "GB")
    for i in range(10):
        add(empty(276), "empty276", str(i))
    add(txt(276, "group B end"), "txt276", "end")
    add(marker("GC"), "M", "GC")
    for h in (36.0, 36.2, 36.4, 36.6):
        idx += 1
        add(img_p(360, h, idx), "img360", str(h))
        add(txt(276, "after image %s" % h), "txt276", "follow")
    add(marker("GD"), "M", "GD")
    for i in range(10):
        add(txt(240, "single text %d" % i), "txt240", str(i))
    add(marker("GE"), "M", "GE")
    for i in range(4):
        add(txt(276, "after160 text %d" % i, after=160), "txt276a160", str(i))
    add(txt(276, "walk end"), "txt276", "final")
    return body, manifest


def gen():
    pitch = int(sys.argv[2]) if len(sys.argv) > 2 else 360
    os.makedirs(OUT, exist_ok=True)
    body, _ = build()
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/>'
           '<w:docGrid w:linePitch="%d"/></w:sectPr></w:body></w:document>' % pitch)
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/media/img.png", PNG)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(body), "paragraphs, linePitch", pitch)


def report(rows):
    _, manifest = build()
    prev = None
    for i, (kind, detail) in enumerate(manifest):
        page, y = rows[i] if i < len(rows) else (None, None)
        adv = ""
        if prev is not None and y is not None and prev[0] == page:
            adv = "adv=%7.3f  px=%7.2f" % (y - prev[1], (y - prev[1]) / 0.75)
        print("p%-3d pg%s y=%8.2f  %-10s %-6s %s"
              % (i, page, -1 if y is None else y, kind, detail, adv))
        prev = (page, y) if y is not None else None


def read():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.ScreenUpdating = False
    d = app.Documents.Open(DOCX, ReadOnly=True)
    rows = []
    try:
        d.Repaginate()
        for p in d.Paragraphs:
            rng = p.Range
            c = d.Range(rng.Start, rng.Start)
            rows.append((c.Information(3), round(c.Information(6), 2)))
    finally:
        d.Close(False)
        app.Quit()
    print("WORD walk:")
    report(rows)


def oxi(envs=""):
    import json
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "gridwalk_oxi.json")
    subprocess.run([GDI, DOCX, os.path.join(tempfile.gettempdir(), "gridwalk"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    # positional recovery: emitted elements only (empty paras emit nothing) —
    # match by text where possible.
    print("OXI dump (line tops per page):")
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        seen = {}
        for e in pg["elements"]:
            key = round(e["y"], 3)
            t = (e.get("text") or "").strip()[:20]
            if key not in seen or (not seen[key][1] and t):
                seen[key] = (e.get("h", 0), t)
        for y in sorted(seen):
            h, t = seen[y]
            print("  pg%d y=%8.3f h=%7.3f %r" % (pg["page"], y, h, t))


if __name__ == "__main__":
    {"gen": gen, "read": read, "oxi": oxi}[sys.argv[1]](*(
        [sys.argv[2]] if sys.argv[1] == "oxi" and len(sys.argv) > 2 else []))
