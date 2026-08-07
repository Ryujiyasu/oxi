# -*- coding: utf-8 -*-
"""Does an inline `w:object` add the text DESCENT on top of its shape height?

S875 derived MIXED H = max(normal_line, obj_h + win_desc + extra) from a 36-arm
sweep whose objects were a PNG inside `w:object`. policies__0016b30b's body line
with a 17pt object measures H = 16.95 in Word (= obj_h, NO descent), and its
object is a real `o:OLEObject ProgID="Equation.DSMT4"`. This probe pins the two
apart at the SAME geometry (Times New Roman 12, line=240 auto).

The line height is read EXACTLY from PDF baselines with a marker sandwich, so
the 0.75pt Information(6) quantization cannot blur obj (17.00) from obj+desc
(19.59):   H(TEST) = g1 + g2 - H(plain),
           g1 = base(TEST) - base(AA marker), g2 = base(ZZ marker) - base(TEST),
with every paragraph forced to before=0 after=0 line=240 auto.

  python _pb_objline2_gen.py gen | bake | read
"""
import base64
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_objline2")
DOCX = os.path.join(OUT, "objline2.docx")
PDF = os.path.join(OUT, "objline2.pdf")

PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAE"
    "hQGAhKmMIQAAAABJRU5ErkJggg==")

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:v="urn:schemas-microsoft-com:vml" '
      'xmlns:o="urn:schemas-microsoft-com:office:office"')

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
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>')

DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" '
         'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
         'Target="styles.xml"/>'
         '<Relationship Id="rImg" '
         'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
         'Target="media/image1.png"/></Relationships>')

FONT = "Times New Roman"
RPR = '<w:rFonts w:ascii="' + FONT + '" w:hAnsi="' + FONT + '"/><w:sz w:val="24"/>'

STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>' + RPR + '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/>'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr>' + RPR + '</w:rPr></w:style></w:styles>')


def obj_run(h, w=10.45, ole=False):
    sid = "s%d" % int(h * 100)
    inner = ('<v:shape id="' + sid + '" type="#_x0000_t75" '
             'style="width:%gpt;height:%gpt">' % (w, h) +
             '<v:imagedata r:id="rImg" o:title=""/></v:shape>')
    if ole:
        inner += ('<o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="' + sid +
                  '" DrawAspect="Content" ObjectID="_1" r:id="rImg"/>')
    return ('<w:r><w:rPr>' + RPR + '</w:rPr>'
            '<w:object w:dxaOrig="%d" w:dyaOrig="%d">' % (int(w * 20), int(h * 20)) +
            inner + '</w:object></w:r>')


def para(inner):
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" '
            'w:lineRule="auto"/><w:rPr>' + RPR + '</w:rPr></w:pPr>' + inner + '</w:p>')


def txt(s):
    return '<w:r><w:rPr>' + RPR + '</w:rPr><w:t xml:space="preserve">' + s + '</w:t></w:r>'


CASES = []
for _oh in (8.0, 17.0, 24.0):
    for _ole in (False,):
        for _mixed in (True, False):
            CASES.append(("%s%02d%s" % ("O" if _ole else "P", int(_oh), "M" if _mixed else "S"),
                          _oh, _ole, _mixed))


def build():
    b = []
    for tag, oh, ole, mixed in CASES:
        b.append(para(txt("AA" + tag)))
        inner = (txt("Tx ") if mixed else "") + obj_run(oh, ole=ole)
        b.append(para(inner))
        b.append(para(txt("ZZ" + tag)))
    b.append(para(txt("END")))
    body = "".join(b) + ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                         '<w:pgMar w:top="720" w:right="1080" w:bottom="720" '
                         'w:left="1080" w:header="708" w:footer="708"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document ' + NS + '><w:body>' + body + '</w:body></w:document>')


def gen():
    os.makedirs(OUT, exist_ok=True)
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/document.xml", build())
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/media/image1.png", PNG)
    print("generated %d cases -> %s" % (len(CASES), os.path.abspath(DOCX)))


def bake():
    import win32com.client as wc
    app = wc.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        d = app.Documents.Open(os.path.abspath(DOCX), ReadOnly=True, AddToRecentFiles=False)
        d.ExportAsFixedFormat(os.path.abspath(PDF), 17)
        d.Close(0)
    finally:
        app.Quit()
    print("baked", os.path.abspath(PDF))


def read():
    import fitz
    doc = fitz.open(PDF)
    lines = []
    for pi, pg in enumerate(doc):
        rows = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    t = "".join(c["c"] for c in s["chars"])
                    if not t.strip():
                        continue
                    y = round(s["origin"][1], 2)
                    k = next((k for k in rows if abs(k - y) <= 0.75), y)
                    rows.setdefault(k, []).append((round(s["origin"][0], 1), t))
        for y in sorted(rows):
            lines.append((pi, y, "".join(t for _, t in sorted(rows[y])).strip()))
    idx = {t: (pi, y) for pi, y, t in lines}
    plain = [lines[i + 1][1] - lines[i][1] for i in range(len(lines) - 1)
             if lines[i][0] == lines[i + 1][0]
             and lines[i][2].startswith("ZZ") and lines[i + 1][2].startswith("AA")]
    mk = round(sum(plain) / len(plain), 3) if plain else 13.8
    desc = 0.216 * 12.0
    print("plain line height (ZZ->AA) = %.3f  (n=%d)" % (mk, len(plain)))
    print("%-8s %6s %7s %7s %8s | %6s %9s" %
          ("case", "obj", "g1", "g2", "H", "obj", "obj+desc"))
    for tag, oh, ole, mixed in CASES:
        a, z = idx.get("AA" + tag), idx.get("ZZ" + tag)
        if not a or not z or a[0] != z[0]:
            print("%-8s %6.2f  (page split / missing)" % (tag, oh))
            continue
        mid = [(pi, y, t) for pi, y, t in lines if pi == a[0] and a[1] < y < z[1]]
        if mid:
            ty = mid[0][1]
            g1, g2 = ty - a[1], z[1] - ty
            h = g1 + g2 - mk
            print("%-8s %6.2f %7.2f %7.2f %8.2f | %6.2f %9.2f" %
                  (tag, oh, g1, g2, h, oh, oh + desc))
        else:
            h = (z[1] - a[1]) - mk
            print("%-8s %6.2f %7s %7s %8.2f | %6.2f %9.2f" %
                  (tag, oh, "-", "-", h, oh, oh + desc))


if __name__ == "__main__":
    {"gen": gen, "bake": bake, "read": read}[sys.argv[1]]()
