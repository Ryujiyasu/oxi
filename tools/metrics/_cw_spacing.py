# -*- coding: utf-8 -*-
"""What does `w:spacing` in a run do to a character's advance, and what does
Word put between a CJK character and a Latin one?

04b88e's 「うちH4年度以降許可債分」 is the line the derived cell budget exposed:
Word fits it in exactly the width it has, Oxi makes it 1.84pt wider and wraps.
The run carries `w:spacing w:val="-20"` at `w:sz w:val="16"`, and the three
pieces disagree separately -- the CJK run by +2.00 over eight characters, the
Latin pair by +0.92, and the CJK/Latin joint by -1.08 because Oxi puts nothing
there at all. Each is its own question, so sweep each.

Advances come from a line long enough never to wrap and left-aligned so Word
never stretches it.

    python _cw_spacing.py           # generate, export through Word, measure
    python _cw_spacing.py --keep    # reuse the export
    python _cw_spacing.py --oxi     # and read Oxi's own advances alongside
"""
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
OUT = os.path.join(REPO, "pipeline_data", "_cw_law")
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MARK = "甲"
# 甲 marks the row; then CJK-CJK pairs, a CJK->Latin joint, Latin-Latin, a
# Latin->CJK joint, and CJK again, all in one line.
TEXT = MARK + "亜亜亜あH4い亜亜"
SPACINGS = [0, -10, -20, -30, -40, -60, 10, 20]     # w:spacing, twentieths of a pt
SIZES = [16, 21]                                     # w:sz, half-points


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


BALANCE = os.environ.get("BALANCE", "") == "1"
ASCII = os.environ.get("ASCII", "")          # "" = same face as the CJK run
COMPAT = os.environ.get("COMPAT", "15")


def build(out, in_cell=True):
    font = "ＭＳ 明朝"
    rows = []
    for sz in SIZES:
        for sp in SPACINGS:
            asc = ASCII or font
            rpr = (f'<w:rFonts w:ascii="{asc}" w:eastAsia="{font}" w:hAnsi="{asc}"/>'
                   + (f'<w:spacing w:val="{sp}"/>' if sp else '')
                   + f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
            para = (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{rpr}</w:rPr></w:pPr>'
                    f'<w:r><w:rPr>{rpr}</w:rPr><w:t>{esc(TEXT)}</w:t></w:r></w:p>')
            if in_cell:
                rows.append(
                    '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>'
                    '<w:tblLayout w:type="fixed"/></w:tblPr>'
                    '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
                    '<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>'
                    + para + '</w:tc></w:tr></w:tbl><w:p/>')
            else:
                rows.append(para)
    sectpr = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
              'w:header="851" w:footer="992" w:gutter="0"/>'
              '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'<w:body>{"".join(rows)}{sectpr}</w:body></w:document>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
              '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
              '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                + ('<w:balanceSingleByteDoubleByteWidth/>' if BALANCE else '')
                + '<w:compat><w:compatSetting w:name="compatibilityMode" '
                f'w:uri="http://schemas.microsoft.com/office/word" w:val="{COMPAT}"/>'
                '</w:compat></w:settings>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    docrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
               '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", docrels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)


def export(docx):
    pdf = docx[:-5] + ".pdf"
    if os.path.exists(pdf) and os.path.getmtime(pdf) > os.path.getmtime(docx):
        return pdf
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx, ReadOnly=True)
    try:
        d.ExportAsFixedFormat(pdf, 17)
    finally:
        d.Close(False)
        app.Quit()
    return pdf


def word_rows(pdf):
    import fitz
    rows = []
    for pg in fitz.open(pdf):
        got = []
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                cs = [c for sp in ln["spans"] for c in sp["chars"]]
                if cs and cs[0]["c"] == MARK:
                    got.append((round(ln["bbox"][1], 1), cs))
        got.sort()
        rows.extend(c for _, c in got)
    return rows


def oxi_rows(docx):
    import json
    exe = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                       "oxi-gdi-renderer.exe")
    with tempfile.TemporaryDirectory() as td:
        dj = os.path.join(td, "d.json")
        subprocess.run([exe, docx, os.path.join(td, "p"), "96", "--dump-layout=" + dj],
                       check=True, capture_output=True)
        d = json.load(open(dj, encoding="utf-8"))
    rows, cur = [], None
    for pg in d["pages"]:
        for e in sorted((e for e in pg["elements"] if e["type"] == "text" and e.get("text")),
                        key=lambda e: (e["y"], e["x"])):
            if e["text"].startswith(MARK):
                cur = []
                rows.append(cur)
            if cur is not None:
                cur.append(e)
    return rows


def main():
    os.makedirs(OUT, exist_ok=True)
    tag = (("_bal" if BALANCE else "")
           + ("_" + ASCII.encode("ascii", "ignore").decode() if ASCII else "")
           + ("_c" + COMPAT if COMPAT != "15" else ""))
    docx = os.path.join(OUT, "cw_spacing%s.docx" % (tag or ""))
    if not ("--keep" in sys.argv and os.path.exists(docx)):
        build(docx)
    rows = word_rows(export(docx))
    arms = [(sz, sp) for sz in SIZES for sp in SPACINGS]
    print(f"{len(rows)} rows read, {len(arms)} expected; "
          f"balanceSBDB={BALANCE} ascii={ASCII or '(same)'} compat={COMPAT}")
    ox = oxi_rows(docx) if "--oxi" in sys.argv else []
    print(f"{'sz':>4}{'w:spacing':>11}{'asked':>8} | {'CJK adv':>9}{'got':>8}"
          f" | {'adv(あ) before H':>17}{'CJK->lat gap':>14}{'/em':>7}"
          + ("   | line width oxi vs Word" if ox else ""))
    for i, (sz, sp) in enumerate(arms):
        if i >= len(rows):
            break
        cs = rows[i]
        # MuPDF invents a space glyph wherever the gap is wide enough, so index
        # by character, never by position: the invented space shifts everything
        # after it and silently relabels every column.
        x = [c["bbox"][0] for c in cs]
        ch = [c["c"] for c in cs]
        cjk = [x[j + 1] - x[j] for j in range(len(cs) - 1)
               if ch[j] == "亜" and ch[j + 1] in "亜あ"]
        cjk = sorted(cjk)[len(cjk) // 2]
        j = ch.index("あ")
        joint = x[j + 1] - x[j]
        extra = ""
        if i < len(ox):
            w_word = x[-1] - x[0] + (cs[-1]["bbox"][2] - cs[-1]["bbox"][0])
            w_oxi = sum(e["w"] for e in ox[i])
            extra = f"   | {w_oxi:7.2f} {w_word:7.2f} {w_oxi - w_word:+6.2f}"
        print(f"{sz / 2:>4.1f}{sp:>11}{sp / 20:>8.2f} | {cjk:>9.3f}"
              f"{cjk - sz / 2.0:>8.3f} | {joint:>17.3f}{joint - cjk:>14.3f}"
              f"{(joint - cjk) / (sz / 2.0):>7.3f}{extra}")


if __name__ == "__main__":
    main()
