# -*- coding: utf-8 -*-
"""A minimal repro for the Latin advance inside a Japanese document.

Six real documents say Word draws a Latin glyph at its exact design em while
Oxi returns a com_tw or 10tw-rounded number that is wrong by as much as 0.38pt
per glyph, with the sign changing character by character so it accumulates
instead of cancelling. This isolates that claim: left-aligned lines, short
enough that nothing wraps and nothing is justified, one Latin run per line
between two CJK markers so the boundary space is measured at the same time.

★It is a FAITHFUL SLICE, not a hand-built minimal docx: only word/document.xml
is replaced and every other part of a real corpus document is kept. A minimal
file puts Word into a degraded font-resolution mode and would have answered for
Cambria under Century's name (see probe_minimal_docx_degraded).

    python _cw_latin_repro.py            # build, export through Word, compare
    python _cw_latin_repro.py --keep     # reuse the existing export
"""
import os
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import _cb_budget as B                                    # noqa: E402
from _cw_latin_adv import (load_tables, nominal_size, oxi_width,  # noqa: E402
                           size_key, family_of)

OUT = os.path.join(B.REPO, "pipeline_data", "_cw_latin_repro")
BASE = "tokyoshugyo"          # a real document that already uses Century + MS Mincho
DOCX = os.path.join(OUT, "latin_repro.docx")

# the characters where com_tw and the design em disagree most, plus a plain word
PAYLOAD = ["0123456789", "wgskiotupand", "APIURLDB", "Washington", "1.5/2.0"]
# ★The third field is w:kern. It is the whole experiment: KERNBREAK gives a
# kern-active run the true em, and LATINEM gives a no-kern run the true em only
# in a document with no CJK body -- so a no-kern Latin run inside a Japanese
# document falls through both and keeps the com_tw width. Each face is measured
# with kerning on and off, everything else held.
ARMS = [("Century", 21, 1), ("Century", 21, 0),
        ("Times New Roman", 21, 1), ("Times New Roman", 21, 0),
        ("ＭＳ Ｐゴシック", 21, 1), ("ＭＳ Ｐゴシック", 21, 0),
        ("Century", 18, 0)]
MARK = "甲"
TAIL = "乙"


def build():
    os.makedirs(OUT, exist_ok=True)
    src = B.docx_for(BASE)
    paras = []
    for ascii_font, sz, kern in ARMS:
        for pay in PAYLOAD:
            rpr = ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="ＭＳ 明朝" '
                   'w:cs="%s"/><w:kern w:val="%d"/><w:sz w:val="%d"/>'
                   '<w:szCs w:val="%d"/>'
                   % (ascii_font, ascii_font, ascii_font,
                      2 if kern else 0, sz, sz))
            paras.append(
                '<w:p><w:pPr><w:jc w:val="left"/><w:rPr>%s</w:rPr></w:pPr>'
                '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s%s%s</w:t>'
                '</w:r></w:p>' % (rpr, rpr, MARK, pay, TAIL))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
            'w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/'
           '2006/main"><w:body>%s%s</w:body></w:document>'
           % ("".join(paras), sect))
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            if it.filename == "word/document.xml":
                continue
            z.writestr(it, zin.read(it.filename))
        z.writestr("word/document.xml", doc)
    zin.close()
    return DOCX


def export(docx, keep):
    pdf = docx[:-5] + ".pdf"
    if keep and os.path.exists(pdf):
        return pdf
    if os.path.exists(pdf):
        os.remove(pdf)
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
    """marker-anchored lines -> [(font, nominal size, [(ch, x_pt)])]"""
    import fitz
    rows = []
    for pg in fitz.open(pdf):
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                chars, font, nom = [], None, None
                for s in ln["spans"]:
                    n = round(s["size"], 2)
                    for c in s["chars"]:
                        # ★MuPDF synthesises a space glyph wherever the pen jumps
                        # a wide gap, and Word's CJK/Latin boundary is exactly
                        # such a jump. Index by the real characters only; the
                        # gap survives in the advance either way.
                        if c["c"].isspace():
                            continue
                        chars.append((c["c"], c["origin"][0]))
                    if font is None or len(s["chars"]) > 2:
                        font, nom = s["font"], n
                if chars and chars[0][0] == MARK:
                    rows.append((font, nom, chars))
    return rows


def oxi_rows(docx):
    import json
    gl = os.path.join(OUT, "oxi_glyphs.json")
    r = subprocess.run([B.GDI, docx, os.path.join(OUT, "png"),
                        "--dump-glyphs=" + gl], capture_output=True)
    if r.returncode != 0:
        sys.exit(r.stderr.decode("utf-8", "replace")[-1500:])
    rows = []
    for pg in json.load(open(gl, encoding="utf-8"))["pages"]:
        byline = {}
        for g in pg["glyphs"]:
            byline.setdefault(round(g["top"], 1), []).append(g)
        for top in sorted(byline):
            gs = sorted(byline[top], key=lambda g: g["x"])
            if gs and gs[0]["char"] == MARK:
                rows.append([(g["char"], g["x"]) for g in gs])
    return rows


def main():
    keep = "--keep" in sys.argv
    docx = build() if not keep or not os.path.exists(DOCX) else DOCX
    pdf = export(docx, keep)
    wrows, orows = word_rows(pdf), oxi_rows(docx)
    metrics, com = load_tables()
    print("Word lines %d / Oxi lines %d" % (len(wrows), len(orows)))

    tot_w = tot_o = tot_d = 0.0
    nglyph = 0
    for (font, drawn, wch), och in zip(wrows, orows):
        nom = nominal_size(drawn)
        wt = "".join(c[0] for c in wch)
        ot = "".join(c[0] for c in och)
        if wt != ot:
            print("  skip (text differs): %r vs %r" % (wt, ot))
            continue
        fam = family_of(font)
        upm, widths = metrics.get(fam, (None, {}))
        ctab = com.get(fam, {}).get(size_key(nom), {})
        print("\n-- %s @ %.1fpt  %s" % (fam, nom, wt))
        print("   %-3s %-9s %-9s %-9s %-9s %s"
              % ("ch", "word", "oxi", "design", "o-w", "via"))
        for i in range(len(wch) - 1):
            ch = wch[i][0]
            wa = wch[i + 1][1] - wch[i][1]
            oa = och[i + 1][1] - och[i][1]
            de = widths.get(str(ord(ch)))
            dpt = None if de is None or not upm else de / upm * drawn
            opt, via = (oxi_width(ch, nom, fam, upm, widths, ctab)
                        if upm else (None, "-"))
            if ch == MARK:
                via = via + " +gap?"
            print("   %-3s %-9.3f %-9.3f %-9s %-9.3f %s"
                  % (ch, wa, oa, "%.3f" % dpt if dpt else "-", oa - wa, via))
            if dpt and ch != MARK:
                tot_w += wa
                tot_o += oa
                tot_d += dpt
                nglyph += 1
    print()
    print("%-18s %-5s %-4s %-10s %-10s %-9s %s"
          % ("face", "size", "kern", "word_run", "oxi_run", "o-w", "text"))
    for (font, drawn, wch), och in zip(wrows, orows):
        wt = "".join(c[0] for c in wch)
        if wt != "".join(c[0] for c in och) or len(wch) < 3:
            continue
        # from the first Latin glyph to the closing CJK marker: the run plus the
        # boundary space, which is what the layout has to reproduce
        wrun = wch[-1][1] - wch[1][1]
        orun = och[-1][1] - och[1][1]
        arm = ARMS[(wrows.index((font, drawn, wch))) // len(PAYLOAD)]
        print("%-18s %-5.1f %-4d %-10.3f %-10.3f %-9.3f %s"
              % (family_of(font), nominal_size(drawn), arm[2], wrun, orun,
                 orun - wrun, wt))
    if nglyph:
        print("\n★ %d glyphs   Word %.2f   Oxi %.2f (%+.2f)   design %.2f (%+.2f)"
              % (nglyph, tot_w, tot_o, tot_o - tot_w, tot_d, tot_d - tot_w))


if __name__ == "__main__":
    main()
