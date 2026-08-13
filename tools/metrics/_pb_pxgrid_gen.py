# -*- coding: utf-8 -*-
"""Does Word snap every Latin line top to the 96dpi device pixel (0.75pt)?

Seven frozen-EN documents put 100% of their paragraph tops on an exact 0.75pt
multiple, including four whose top margin is NOT a pixel multiple (1417tw =
70.85pt renders its first line at 70.50 = 94px).  Two boundary cases pin the
accumulation shape:

  policies__000f7115  Arial 12 x1.15   h = 15.87 = 21.16px
      1 line  -> 21px  = 15.75      (its list items advance 15.75 each)
      4 lines -> 85px  = 63.75      (measured 661.50 -> 725.25)
      6 lines -> 127px = 95.25      (measured 534.75 -> 630.00)
  reports__00377a16   TNR 12 x1.5     h = 20.70 = 27.60px
      7 gaps  -> 193px = 144.75     (measured 174.75 -> 319.50)

★RESOLVED by _pb_pxrun_gen.py (run it FIRST): the cursor is NOT snapped.  A
20-paragraph run separates the two readings, and all 8 combos came out EXACT --
Word accumulates the true h and only the REPORTED/painted position is rounded
to the pixel.  So Info6 in a Latin document carries +-0.375pt of reporting
noise: never do sub-point drift arithmetic on a single Info6 advance, only on a
long run (the Latin twin of the CJK Info6-centering trap).

This probe stays as the per-count evidence (font x size x multiplier x line
count, each count forced with <w:br/> so no wrap width is involved); its
"cumulative snap" fit column is a rounding coincidence at short counts and is
kept only to show WHY the longer run was needed.

  python _pb_pxgrid_gen.py gen
  python _pb_pxgrid_gen.py read     # Word COM truth vs the model
  python _pb_pxgrid_gen.py oxi      # what Oxi's cursor does today
"""
import os
import re
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_pxgrid")
DOCX = os.path.join(OUT, "pxgrid.docx")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
PX = 0.75

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/>'
          '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '</w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/></w:style></w:styles>')

FONTS = ["Arial", "Times New Roman", "Calibri"]
SIZES = [20, 22, 24, 28]          # half-points: 10 / 11 / 12 / 14pt
MULTS = [240, 259, 276, 300, 360, 480]
COUNTS = [1, 2, 3, 4, 5, 6]


def combos():
    return [(f, sz, ml) for f in FONTS for sz in SIZES for ml in MULTS]


def para(tag, font, sz, ml, n):
    rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/><w:sz w:val="%d"/>'
           '<w:szCs w:val="%d"/></w:rPr>' % (font, font, font, sz, sz))
    body = ("<w:br/>".join("<w:t>%s</w:t>" % (tag if i == 0 else "x") for i in range(n)))
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="auto"/>'
            '<w:widowControl w:val="0"/>%s</w:pPr><w:r>%s%s</w:r></w:p>'
            % (ml, rpr, rpr, body))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ci, (f, sz, ml) in enumerate(combos()):
        for n in COUNTS:
            body.append(para("C%03dN%d" % (ci, n), f, sz, ml, n))
    body.append('<w:p/>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           '><w:body>' + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>')
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(combos()) * len(COUNTS), "paragraphs")


def word_truth():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(DOCX, ReadOnly=True)
    rows = {}
    try:
        d.Repaginate()
        n = d.Paragraphs.Count
        seq = []
        for i in range(1, n + 1):
            rng = d.Paragraphs(i).Range
            t = rng.Text
            m = re.match(r"(C\d+N\d)", t)
            if not m:
                continue
            c = d.Range(rng.Start, rng.Start)
            seq.append((m.group(1), c.Information(3), round(c.Information(6), 2)))
        for (tag, pg, y), (_t2, pg2, y2) in zip(seq, seq[1:]):
            rows[tag] = (pg, y, (y2 - y) if pg2 == pg else None)
    finally:
        d.Close(False)
        app.Quit()
    return rows


def oxi_heights():
    """Oxi's own per-paragraph line box height, from the widow trace."""
    env = dict(os.environ, OXI_DUMP_WIDOW="1")
    p = subprocess.run([GDI, DOCX, os.path.join(OUT, "px"),
                        "--dump-layout=" + os.path.join(OUT, "px.json")],
                       capture_output=True, env=env, text=True, errors="replace")
    out = {}
    for line in p.stderr.splitlines():
        m = re.search(r"cursor_y=([\d.]+) lh=([\d.]+).*text=\"(C\d+N\d)", line)
        if m:
            out[m.group(3)] = (float(m.group(1)), float(m.group(2)))
    return out


def report():
    """Fit h per combo from Word's own advances -- no Oxi metric assumed.

    Under the cumulative-snap model adv(n) = round(n*h/PX)*PX, each measured
    advance constrains h to a half-open interval; a non-empty intersection
    over n = 1..6 means one single h explains every count.  The rival model
    (a per-line height quantised once, adv(n) = n*adv(1)) is reported next to
    it, and Oxi's own line box is printed where it is available.
    """
    word = word_truth()
    oxi = oxi_heights()
    print("%-4s %-16s %4s %4s | %-28s %-9s %s"
          % ("c", "font", "sz", "mult", "h interval (cumulative snap)", "oxi_lh", "advances"))
    fit = flat = neither = 0
    for ci, (f, sz, ml) in enumerate(combos()):
        advs = []
        for n in COUNTS:
            wr = word.get("C%03dN%d" % (ci, n))
            advs.append(None if not wr or wr[2] is None else wr[2])
        pairs = [(n, a) for n, a in zip(COUNTS, advs) if a is not None]
        if len(pairs) < 4:
            continue
        lo, hi = 0.0, 1e9
        for n, a in pairs:
            lo = max(lo, (a / PX - 0.5) * PX / n)
            hi = min(hi, (a / PX + 0.5) * PX / n)
        a1 = dict(pairs).get(1)
        is_flat = a1 is not None and all(abs(a - n * a1) < 0.03 for n, a in pairs)
        ok = lo < hi
        if ok:
            fit += 1
        if is_flat:
            flat += 1
        if not ok:
            neither += 1
        lh = oxi.get("C%03dN1" % ci, (0, 0))[1]
        tag = ("FIT" if ok else "NO-FIT") + ("/flat" if is_flat else "")
        if not ok or ci % 6 == 0:
            print("%-4d %-16s %4.1f %4d | [%7.3f, %7.3f) %-6s %9.3f  %s"
                  % (ci, f, sz / 2.0, ml, lo, hi, tag, lh,
                     " ".join("-" if a is None else "%.2f" % a for a in advs)))
    print(f"\ncombos fitting ONE h under cumulative snap: {fit}"
          f"   also flat (adv(n)=n*adv(1)): {flat}   no fit: {neither}")


if __name__ == "__main__":
    {"gen": gen, "read": report, "oxi": lambda: print(oxi_heights())}[sys.argv[1]]()
