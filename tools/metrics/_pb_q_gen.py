# -*- coding: utf-8 -*-
"""Task Q — the plain page-bottom last-line fit threshold for single-spacing Latin.

reference__00525b22's p15 boundary: a 3-line Calibri-11 single-spacing paragraph
whose LAST line ("2015.") has bbox bottom 733.656 <= content bottom 734.400 and
Word KEEPS it, while Oxi's widow lookahead measures the last line by its hhea box
(13.428) from the line top -> 734.821 > 734.400 -> whole-moves the paragraph.

The question: does Word fit the last line by
  (A) baseline + FONT descent (= hhea box, Oxi's model), or
  (B) baseline + GLYPH INK descent (content-dependent), or
  (C) something else?

Discriminator = the SAME line with a SHALLOW last glyph set (no descender:
"2015") vs a DEEP one (descenders: "aplogy gjy"). If Word's flip point moves
with the descender, the rule is ink-based (B); if not, it is font-box (A).

Faithful to the target: no-type docGrid linePitch=360, Letter, top/bottom 1152tw,
header/footer 720. Filler single-line paras place the target's line 1 near the
bottom; the bottom margin is swept in 1-twip steps to pin the flip.

  python _pb_q_gen.py gen      -> pipeline_data/_pb_q/
  python _pb_q_gen.py measure  -> Word COM: per-line page + Information(6) y
"""
import os, sys, json, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_q")

DOC_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
           '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {DOC_NS}><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
          '<w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:rPr><w:sz w:val="24"/></w:rPr></w:style></w:styles>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings {DOC_NS}><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '</w:compat></w:settings>')

# last-line content variants (line 3 of the target paragraph). Each is short so it
# is exactly the 3rd wrapped line. SHALLOW = no descenders; DEEP = descenders.
SHALLOW = "2015."     # digits + period, ink descent ~0.25em
DEEP = "apology gjy." # p g j y descenders


def _rpr(font, sz):
    return f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz}"/>'


def filler(i, font, sz):
    r = _rpr(font, sz)
    return (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t>Filler line {i}.</w:t></w:r></w:p>')


def target(font, sz, spacing_attr, last):
    # a sentence that wraps to 2 full lines then `last` as the 3rd line.
    r = _rpr(font, sz)
    body = ("This target paragraph is deliberately written at length so that it "
            "wraps to exactly three separate lines within the text column when set "
            "at the eleven point body size that is used throughout this document, "
            "finally ending in ")
    txt = body + last
    ppr = f'<w:pPr>{spacing_attr}<w:rPr>{r}</w:rPr></w:pPr>'
    return f'<w:p>{ppr}<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{txt}</w:t></w:r></w:p>'


def build(nfill, bottom, font, sz, spacing_attr, last):
    pgsz = '<w:pgSz w:w="12240" w:h="15840" w:code="1"/>'
    mar = (f'<w:pgMar w:top="1152" w:right="1296" w:bottom="{bottom}" '
           f'w:left="1296" w:header="720" w:footer="720" w:gutter="0"/>')
    grid = '<w:docGrid w:linePitch="360"/>'
    fillers = "".join(filler(i + 1, font, sz) for i in range(nfill))
    tgt = target(font, sz, spacing_attr, last)
    sect = f'<w:sectPr>{pgsz}{mar}{grid}</w:sectPr>'
    body = fillers + tgt + f'<w:p><w:pPr><w:sectPr>{pgsz}{mar}{grid}</w:sectPr></w:pPr></w:p>'
    # sectPr must be the LAST paragraph's or the body's; put it in body via a
    # trailing empty para to avoid altering the target.
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {DOC_NS}><w:body>{fillers}{tgt}'
           f'<w:p><w:pPr>{sect}</w:pPr></w:p></w:body></w:document>')
    return doc


# Cases. Calibri 11pt single spacing is the target regime. nfill chosen so the
# target's line 1 lands ~700pt; sweep bottom in 1-twip steps around the flip.
# content bottom = (15840 - bottom)/20; want ~734 -> bottom ~ 1152..1200.
CASES = []
SINGLE = ''            # no <w:spacing> = single
X115 = '<w:spacing w:line="276" w:lineRule="auto"/>'
X200 = '<w:spacing w:line="480" w:lineRule="auto"/>'

# Bottom-margin sweep (content bottom (15840-bottom)/20 = 733..722pt) with the
# target's line 3 near the page bottom. nfill per regime from calibration.
# Calibri 11 single: shallow vs deep last line = the ink-vs-fontbox discriminator.
for last in ("shallow", "deep"):
    lc = SHALLOW if last == "shallow" else DEEP
    for bottom in range(1180, 1361, 4):   # 0.2pt steps, ~9pt window
        CASES.append((47, bottom, "Calibri", 22, SINGLE, lc, f"cal_{last}_b{bottom}"))
# Arial 11 single, shallow AND deep: font-independence of the model.
for last in ("shallow", "deep"):
    lc = SHALLOW if last == "shallow" else DEEP
    for bottom in range(1180, 1361, 4):
        CASES.append((50, bottom, "Arial", 22, SINGLE, lc, f"ari_{last}_b{bottom}"))
# TNR 12 single, shallow vs deep: the report claims TNR uses hhea (no descender effect).
for last in ("shallow", "deep"):
    lc = SHALLOW if last == "shallow" else DEEP
    for bottom in range(1180, 1361, 4):
        CASES.append((46, bottom, "Times New Roman", 24, SINGLE, lc, f"tnr_{last}_b{bottom}"))
# Calibri 11 x1.15, shallow: multiple spacing -> report claims natural/hhea (leading hangs).
for bottom in range(1180, 1361, 4):
    CASES.append((47, bottom, "Calibri", 22, X115, SHALLOW, f"cal115_shallow_b{bottom}"))


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for nfill, bottom, font, sz, sp, last, nm in CASES:
        parts = {
            "[Content_Types].xml": CT, "_rels/.rels": RELS,
            "word/document.xml": build(nfill, bottom, font, sz, sp, last),
            "word/_rels/document.xml.rels": DOCRELS,
            "word/styles.xml": STYLES, "word/settings.xml": SETTINGS,
        }
        with zipfile.ZipFile(os.path.join(OUTDIR, nm + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
            for n, d in parts.items():
                z.writestr(n, d)
    print("generated", len(CASES))


def measure():
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    out = {}
    try:
        for nfill, bottom, font, sz, sp, last, nm in CASES:
            p = os.path.abspath(os.path.join(OUTDIR, nm + ".docx"))
            doc = word.Documents.Open(p, ReadOnly=True)
            try:
                # the target is paragraph nfill+1. Its start / mid / end pages.
                tpar = doc.Paragraphs(nfill + 1).Range
                s = doc.Range(tpar.Start, tpar.Start)
                e = doc.Range(tpar.End - 1, tpar.End - 1)
                pg_start = int(s.Information(3))
                pg_end = int(e.Information(3))
                y_start = float(s.Information(6))
                y_end = float(e.Information(6))
                out[nm] = dict(bottom=bottom, pg_start=pg_start, pg_end=pg_end,
                               y_start=round(y_start, 2), y_end=round(y_end, 2))
                print(nm, out[nm], flush=True)
            finally:
                doc.Close(False)
    finally:
        word.Quit()
    with open(os.path.join(OUTDIR, "_measure.json"), "w") as f:
        json.dump(out, f, indent=1)
    print("wrote", len(out))


if __name__ == "__main__":
    if sys.argv[1:] == ["gen"]:
        gen()
    elif sys.argv[1:] == ["measure"]:
        measure()
    else:
        print(__doc__)
