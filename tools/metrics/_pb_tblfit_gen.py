# -*- coding: utf-8 -*-
"""How does Word shrink an AUTO-layout table whose gridCols exceed the page?

reports__001f1397 (no tblPr at all, grid 1228/3416/2385/1987 = 9016tw into an
8504tw column) is laid out by Word as 1227/3182/2286/1809 — the 512tw excess is
NOT taken proportionally (c0 keeps its full width, c3 loses 178). Its c3 ends up
at ~the width of its longest word ("Conservative"), which is what a content-
driven autofit predicts. Oxi shrinks proportionally, so its c1 stays wide enough
to fit one extra word per line, the row measures 20pt short, and the doc
paginates -1 (score 0.5698; 58/619 corpus docs carry an over-wide auto table).

Each arm is one page holding one 4-column table with a KNOWN grid and known cell
contents, so the min (longest word) and max (unwrapped paragraph) width of every
column is computable. Word's resulting column boundaries are read from the PDF:
each cell opens with a marker glyph, so cell_left = marker_x - inset.

Arms discriminate:
  base      excess with all columns holding wrappable text (slack everywhere)
  minwide   one column whose longest WORD exceeds its proportional share
  minall    every column at its minimum -> excess cannot be absorbed
  empty     one empty column (min = 0)
  sweep     the same grid at 3 excess magnitudes -> linear or clamped?
  nowrap    a column whose content is one long unbreakable token

  python _pb_tblfit_gen.py gen
  python _pb_tblfit_gen.py pdf     # Word truth (ExportAsFixedFormat + fitz)
  python _pb_tblfit_gen.py oxi     # same columns from --dump-layout
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_autofit")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

# A4 portrait, 1440tw left/right margins -> 11906 - 2880 = 9026tw available
PGW, MARG = 11906, 1440
AVAIL = PGW - 2 * MARG

# (name, gridCols, [cell texts])  — 4 columns everywhere so the readout is uniform
SHORT = "aa bb cc dd ee ff gg hh ii jj kk ll mm nn oo pp"
LONGW = "Conservative Liberal Democrat Coalition"
HUGEW = "Incomprehensibilities"          # one 21-char token, no break opportunity
ARMS = [
    # specimen-shaped: c0 narrow with short tokens, c3 holding long words
    ("A_spec",    [1228, 3416, 2385, 1987], ["20010/11", SHORT, "1.5%", LONGW]),
    # same grid, but every column has small min (short tokens everywhere)
    ("B_allshort", [1228, 3416, 2385, 1987], ["aa", SHORT, "bb", "cc dd ee"]),
    # one column whose longest token is wide
    ("C_minwide", [1228, 3416, 2385, 1987], ["20010/11", SHORT, HUGEW, LONGW]),
    # an empty column
    ("D_empty",   [1228, 3416, 2385, 1987], ["20010/11", SHORT, "", LONGW]),
    # excess sweep on a uniform grid: 200 / 1000 / 3000tw over
    ("E_over200", [(AVAIL + 200) // 4] * 4, ["aa", SHORT, "bb", LONGW]),
    ("F_over1000", [(AVAIL + 1000) // 4] * 4, ["aa", SHORT, "bb", LONGW]),
    ("G_over3000", [(AVAIL + 3000) // 4] * 4, ["aa", SHORT, "bb", LONGW]),
    # every column holds one long token -> minimums alone exceed the page
    ("H_minall",  [(AVAIL + 3000) // 4] * 4, [HUGEW] * 4),
    # control: grid FITS (no shrink) — the readout must return the grid verbatim
    ("I_fits",    [1228, 2000, 2385, 1987], ["20010/11", SHORT, "1.5%", LONGW]),
    # ★uklocalspending counterexample (2026-08-15): its 14018tw grid overflows a
    # 13778tw landscape column and Word renders it at FULL width — no shrink at
    # all. That table differs from the arms above in exactly two ways, so both
    # get an arm: an EXPLICIT `<w:tblW w:w="0" w:type="auto"/>` and a 15tw
    # tblCellMar (vs the 108tw the other arms inherit per-cell).
    ("J_tblw_auto", [(AVAIL + 1000) // 4] * 4, ["aa", SHORT, "bb", LONGW]),
    ("K_cellmar15", [(AVAIL + 1000) // 4] * 4, ["aa", SHORT, "bb", LONGW]),
    ("L_both",      [(AVAIL + 1000) // 4] * 4, ["aa", SHORT, "bb", LONGW]),
    # ★the remaining difference: uklocalspending's CELLS declare
    # `<w:tcW w:w="0" w:type="auto"/>` (no preferred width at all), where every
    # arm above declares a dxa preferred width equal to its gridCol.
    ("M_tcw_auto",  [(AVAIL + 1000) // 4] * 4, ["aa", SHORT, "bb", LONGW]),
    ("N_ukls",      [(AVAIL + 1000) // 4] * 4, ["aa", SHORT, "bb", LONGW]),
]

# arm name -> (tblPr xml, per-cell tcMar twips, tcW type)
VARIANT = {
    "J_tblw_auto": ('<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>', 108, "dxa"),
    "K_cellmar15": ("", 15, "dxa"),
    "L_both": ('<w:tblPr><w:tblW w:w="0" w:type="auto"/>'
               '<w:tblCellMar><w:left w:w="15" w:type="dxa"/>'
               '<w:right w:w="15" w:type="dxa"/></w:tblCellMar></w:tblPr>', 15, "dxa"),
    "M_tcw_auto": ("", 108, "auto"),
    "N_ukls": ('<w:tblPr><w:tblW w:w="0" w:type="auto"/>'
               '<w:tblCellMar><w:left w:w="15" w:type="dxa"/>'
               '<w:right w:w="15" w:type="dxa"/></w:tblCellMar></w:tblPr>', 15, "auto"),
}

MARK = "■"          # cell-left marker glyph (BLACK SQUARE, Tahoma has it)
FS = 24                  # half-points -> 12pt Tahoma, as the specimen


def docx():
    return os.path.join(OUT, "autofit.docx")


def rpr():
    return ('<w:rPr><w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma" w:cs="Tahoma"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (FS, FS))


def para(text, pbb=False):
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:jc w:val="left"/>%s</w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % ("<w:pageBreakBefore/>" if pbb else "", rpr(), rpr(), text))


def cell(text, w, mar=108, wtype="dxa"):
    # tcMar 108/108 as the specimen; tcW dxa unless the arm asks for auto
    return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="%s"/>'
            '<w:tcBorders><w:top w:val="single" w:sz="4" w:color="000000"/>'
            '<w:left w:val="single" w:sz="4" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="4" w:color="000000"/>'
            '<w:right w:val="single" w:sz="4" w:color="000000"/></w:tcBorders>'
            '<w:tcMar><w:left w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/></w:tcMar>'
            '</w:tcPr>%s</w:tc>' % (0 if wtype == "auto" else w, wtype, mar, mar,
                                     para(MARK + text)))


def table(cols, texts, name=""):
    # NO tblPr at all by default — the specimen's shape (auto layout, no tblW)
    tblpr, mar, wtype = VARIANT.get(name, ("", 108, "dxa"))
    grid = "".join('<w:gridCol w:w="%d"/>' % c for c in cols)
    tcs = "".join(cell(t, w, mar, wtype) for t, w in zip(texts, cols))
    return ("<w:tbl>%s<w:tblGrid>%s</w:tblGrid><w:tr>%s</w:tr></w:tbl>"
            % (tblpr, grid, tcs))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (name, cols, texts) in enumerate(ARMS):
        body.append(para("M%02d %s" % (ai, name), pbb=ai > 0))
        body.append(table(cols, texts, name))
        body.append(para("E%02d" % ai))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="%d" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="%d" w:bottom="1440" w:left="%d" '
           'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
           % (PGW, MARG, MARG))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma" w:cs="Tahoma"/><w:sz w:val="24"/>'
              "</w:rPr></w:rPrDefault>"
              '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
              "</w:pPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style></w:styles>')
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms | avail =", AVAIL, "tw")


def report(who, per_arm):
    """per_arm[ai] = [cell_left_pt, ...] (4 entries) + table_right_pt"""
    print("== %s  (avail %dtw = %.2fpt)" % (who, AVAIL, AVAIL / 20.0))
    print("%-11s %-28s %-28s %s" % ("arm", "grid (tw)", "measured (tw)", "delta"))
    for ai, (name, cols, _t) in enumerate(ARMS):
        xs = per_arm.get(ai)
        if not xs or len(xs) < len(cols) + 1:
            print("%-11s MISSING" % name)
            continue
        meas = [int(round((xs[i + 1] - xs[i]) * 20.0)) for i in range(len(cols))]
        delta = [m - g for m, g in zip(meas, cols)]
        print("%-11s %-28s %-28s %s  (sum %d vs %d)"
              % (name, " ".join("%5d" % c for c in cols),
                 " ".join("%5d" % m for m in meas),
                 " ".join("%+5d" % d for d in delta), sum(meas), sum(cols)))


def _cells_from_xs(xs, inset_pt=5.4):
    """marker x -> cell left (subtract the 108tw inset + half the border)."""
    return [x - inset_pt for x in xs]


def _vertical_border_xs(pg):
    """Column boundaries from the table's VERTICAL border strokes.

    The marker-glyph readout was 142tw off on column 0 (the first cell's left
    inset is not the same distance the table is outdented by), so the boundaries
    are taken from the drawn borders instead: every near-vertical stroke's x,
    deduplicated at 0.6pt.
    """
    xs = []
    for dr in pg.get_drawings():
        r = dr["rect"]
        if r.height > 4.0 and r.width <= 2.5:      # a vertical rule
            xs.append(round((r.x0 + r.x1) / 2.0, 2))
    xs.sort()
    out = []
    for x in xs:
        if not out or x - out[-1] > 0.6:
            out.append(x)
    return out


def pdf():
    import fitz
    import win32com.client as w
    out = docx().replace(".docx", ".pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx(), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(out, 17)
    finally:
        d.Close(False)
        app.Quit()
    doc = fitz.open(out)
    per = {}
    for ai in range(len(ARMS)):
        if ai >= doc.page_count:
            break
        xs = _vertical_border_xs(doc[ai])
        if len(xs) >= 2:
            per[ai] = xs
    report("WORD", per)


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "autofit_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "af"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    per = {}
    for ai in range(len(ARMS)):
        if ai >= len(pages):
            break
        # ★borders, not the marker glyph: the marker readout assumes a fixed
        # 5.4pt inset and reports a 15tw-margin arm ~93tw wide (the layout is
        # right; the reporter was not). Same basis as the Word side.
        bx = sorted({round(e["x"], 2) for e in pages[ai]["elements"]
                     if e.get("type") == "border"})
        ded = []
        for x in bx:
            if not ded or x - ded[-1] > 0.6:
                ded.append(x)
        if len(ded) >= 2:
            per[ai] = ded
    report("OXI " + (envs or "(default)"), per)


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
