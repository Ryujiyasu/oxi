# -*- coding: utf-8 -*-
"""Whose spacing-BEFORE does a run-in (`w:specVanish`) heading's merged
paragraph use — the HEADING's, or the continuation paragraph's?

Last blocker for S1189 (see docs/spec/table_border_draw_width_2026_08_21.md).
technical__00501ca3 stacks

    <w:p><w:pPr><w:spacing w:after="200"/></w:pPr>   (Amended 2003)
    <w:p><w:pPr><w:pStyle Heading4/>
         <w:rPr><w:vanish/><w:specVanish/></w:rPr></w:pPr>  T.1.1. Repeatability.
    <w:p><w:pPr><w:pStyle BodyTextIndent/>
         <w:spacing w:before="240"/></w:pPr>          - When multiple tests ...

Word merges the heading into the body paragraph (its PDF line reads
"T.1.1. Repeatability. - When multiple tests are cond...") and advances
    649.19 -> 670.67 = 21.48 = line 11.48 (Times New Roman 10 hhea) + after 10.0
i.e. the continuation paragraph's DIRECT `before=240` (12pt) is NOT applied, and
Heading4 carries no spacing of its own. Oxi advances 23.50 = 11.48 + 12.02, so
it takes the FOLLOWING paragraph's before — S784 hands the merged paragraph the
following para's properties wholesale, which is right for the mark but appears
wrong for space-before.

Arms isolate exactly that: whose `before` shows up in the gap.

  python tools/metrics/_pb_specvanish_gen.py --measure --oxi
"""
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = Path(os.environ.get("OXI_SCRATCH", tempfile.gettempdir())) / "pb_specvanish.docx"
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def mark(i, side):
    return "ZMARK%s%02dZ" % (side, i)


def run(t):
    return ('<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r>' % t)


def sp(before=None, after=None):
    if before is None and after is None:
        return ""
    a = ' w:after="%d"' % after if after is not None else ""
    b = ' w:before="%d"' % before if before is not None else ""
    return "<w:spacing%s%s/>" % (b, a)


def para(text, before=None, after=None, specvanish=False):
    rpr = "<w:rPr><w:vanish/><w:specVanish/></w:rPr>" if specvanish else ""
    return ("<w:p><w:pPr>" + sp(before, after) + rpr + "</w:pPr>" + run(text) + "</w:p>")


# (name, head_before, head_after, body_before)  — head is the specVanish run-in
ARMS = [
    ("runin_body240", None, None, 240),
    ("runin_body0", None, None, None),
    ("runin_head360_body240", 360, None, 240),
    ("runin_head_after360", None, 360, 240),
    ("norunin_body240", None, None, 240),      # control: no specVanish
    ("runin_body480", None, None, 480),
]
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def build():
    body = []
    for i, (name, hb, ha, bb) in enumerate(ARMS):
        brk = '<w:p><w:pPr><w:pageBreakBefore/></w:pPr>' + run(mark(i, "A")) + "</w:p>" if i \
            else "<w:p>" + run(mark(i, "A")) + "</w:p>"
        body.append(brk)
        # the anchor paragraph: after=200 (10pt), like the real doc
        body.append(para("ANCHORLINE", after=200))
        body.append(para("HEADRUNIN.", before=hb, after=ha,
                         specvanish=(name != "norunin_body240")))
        body.append(para("BODYTEXTHERE", before=bb))
        body.append("<w:p>" + run(mark(i, "B")) + "</w:p>")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document %s><w:body>%s%s</w:body></w:document>' % (NS, "".join(body), SECT))
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
    print("wrote", OUT)
    return OUT


def word_lines(path):
    pdf = Path(tempfile.gettempdir()) / (path.stem + ".truth.pdf")
    if not pdf.exists() or "--reexport" in sys.argv:
        import win32com.client as win32
        w = win32.DispatchEx("Word.Application")
        w.Visible = False
        try:
            d = w.Documents.Open(str(path), ReadOnly=True)
            d.ExportAsFixedFormat(str(pdf), 17)
            d.Close(False)
        finally:
            w.Quit()
    import fitz
    doc = fitz.open(pdf)
    out = []
    for pi in range(doc.page_count):
        for blk in doc[pi].get_text("dict")["blocks"]:
            if blk.get("type", 0) != 0:
                continue
            for ln in blk.get("lines", []):
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if t:
                    out.append((pi * 10000 + min(s["bbox"][1] for s in ln["spans"]), t))
    return out


def oxi_lines(path):
    exe = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
    tmp = Path(tempfile.mkdtemp())
    dump = tmp / "d.json"
    subprocess.run([str(exe), str(path), str(tmp / "p"), "110", "--dump-layout=%s" % dump],
                   check=True, capture_output=True)
    import json
    d = json.load(open(dump, encoding="utf-8"))
    rows = {}
    for pi, pg in enumerate(d["pages"]):
        for e in pg["elements"]:
            if e.get("type") == "text" and (e.get("text") or "").strip():
                rows.setdefault(pi * 10000 + round(e.get("y", 0.0), 2), []).append(
                    (e.get("x", 0.0), e["text"]))
    return [(y, "".join(t for _, t in sorted(v))) for y, v in rows.items()]


def summarize(rows, tag):
    rows.sort()
    print("--- %s ---" % tag)
    res = {}
    for i, (name, hb, ha, bb) in enumerate(ARMS):
        def find(sub):
            return [y for y, t in rows if sub in t.replace(" ", "")]
        anc = find("ANCHORLINE")
        head = find("HEADRUNIN")
        bodyl = find("BODYTEXTHERE")
        a = [y for y in anc if any(abs(y - x) < 400 and y > x for x in find(mark(i, "A")))]
        if not a or not head:
            print("  %-22s MISSING" % name)
            continue
        anchor_y = min(a)
        # the run-in line is where HEADRUNIN sits; merged iff BODYTEXTHERE shares it
        hy = min(y for y in head if y > anchor_y)
        by = min(y for y in bodyl if y >= anchor_y) if bodyl else None
        merged = by is not None and abs(by - hy) < 0.5
        print("  %-22s anchor %8.2f  runin %8.2f  gap %6.2f  merged=%s"
              % (name, anchor_y, hy, hy - anchor_y, merged))
        res[name] = (hy - anchor_y, merged)
    return res


if __name__ == "__main__":
    p = build()
    w = summarize(word_lines(p), "WORD") if "--measure" in sys.argv else None
    o = summarize(oxi_lines(p), "OXI") if "--oxi" in sys.argv else None
    if w and o:
        print("--- DIFF (oxi - word) ---")
        for arm in ARMS:
            n = arm[0]
            if n in w and n in o:
                print("  %-22s d_gap %+7.2f   (word %.2f / oxi %.2f)  merged w=%s o=%s"
                      % (n, o[n][0] - w[n][0], w[n][0], o[n][0], w[n][1], o[n][1]))
