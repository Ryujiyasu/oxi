# -*- coding: utf-8 -*-
"""Derive Word's BORDER CONFLICT RESOLUTION at a shared horizontal edge.

The blocker for S1187 (see docs/spec/table_border_draw_width_2026_08_21.md):
`technical__00501ca3` declares `insideH single sz6` on the table AND `double
sz4` on many cell tops/bottoms. Word draws mostly the SINGLE; Oxi fires the
cell-declared double 190x where Word draws ~5 doubles per page. Oxi has no model
of which declaration wins on the edge two rows share — and that same missing
rule is what S865's `thick_fixed` carve-out stands in for.

Each arm is one 4-row 2-column table on its own page. The table declares
`insideH`; row 2's cells may declare a `top` and row 1's cells a `bottom`. The
PDF then says which one Word actually drew (thickness, and whether it is a
double's 2-rect pair) and how the row pitch responded — the pitch is the part
the layout engine has to get right.

Arms also separate two questions Oxi currently conflates:
  * WHICH border wins on the shared edge (style/width weight)
  * WHETHER one cell's declaration spans the whole row or only its own column
    (`one_cell_*`: only the FIRST cell of row 2 declares it)

  python tools/metrics/_pb_bconflict_gen.py            # write the probe docx
  python tools/metrics/_pb_bconflict_gen.py --measure  # + Word PDF truth
  python tools/metrics/_pb_bconflict_gen.py --oxi      # + Oxi's answer
"""
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = Path(os.environ.get("OXI_SCRATCH", tempfile.gettempdir())) / "pb_bconflict.docx"

NS = ' '.join([
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
])
N_ROWS = 4
EDGE_ROW = 1  # the shared edge under test is between row 1 and row 2 (0-based)


def bd(tag, spec):
    """One border element. spec = (style, sz) or ('nil', 0) or None."""
    if spec is None:
        return ""
    style, sz = spec
    if style == "nil":
        return '<w:%s w:val="nil"/>' % tag
    return '<w:%s w:val="%s" w:sz="%d" w:space="0" w:color="auto"/>' % (tag, style, sz)


def mark(i, side):
    return "ZMARK%s%02dZ" % (side, i)


def txt(s, sz=20):
    return ('<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r>' % (sz, sz, s))


def cell(text, top=None, bottom=None):
    tb = ""
    if top is not None or bottom is not None:
        tb = "<w:tcBorders>" + bd("top", top) + bd("bottom", bottom) + "</w:tcBorders>"
    return ('<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>' + tb + "</w:tcPr>"
            '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            + txt(text) + "</w:p></w:tc>")


def tbl(inside_h, r1_bot=None, r2_top=None, first_cell_only=False):
    edges = "".join(bd(k, ("single", 4)) for k in ("top", "left", "bottom", "right"))
    edges += bd("insideH", inside_h) + bd("insideV", ("single", 4))
    rows = []
    for k in range(N_ROWS):
        # row EDGE_ROW's bottom and row EDGE_ROW+1's top form the edge under test
        bot = r1_bot if k == EDGE_ROW else None
        top = r2_top if k == EDGE_ROW + 1 else None
        c0 = cell("R%d" % k, top=top, bottom=bot)
        if first_cell_only:
            c1 = cell("C%d" % k)          # second cell declares nothing
        else:
            c1 = cell("C%d" % k, top=top, bottom=bot)
        rows.append("<w:tr>" + c0 + c1 + "</w:tr>")
    return ('<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>'
            '<w:tblLayout w:type="fixed"/><w:tblBorders>' + edges + "</w:tblBorders></w:tblPr>"
            '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
            + "".join(rows) + "</w:tbl>")


S6 = ("single", 6)
S4 = ("single", 4)
S12 = ("single", 12)
S24 = ("single", 24)
D4 = ("double", 4)
NIL = ("nil", 0)

# (name, insideH, row1_bottom, row2_top, first_cell_only)
ARMS = [
    ("base_ih6", S6, None, None, False),
    ("ih6_top_d4", S6, None, D4, False),
    ("ih6_bot_d4", S6, D4, None, False),
    ("ih6_both_d4", S6, D4, D4, False),
    ("ih6_top_nil", S6, None, NIL, False),
    ("ih6_both_nil", S6, NIL, NIL, False),
    ("ih6_top_s12", S6, None, S12, False),
    ("ih12_top_s6", S12, None, S6, False),
    ("ih24_top_s4", S24, None, S4, False),
    ("ihNone_top_d4", None, None, D4, False),
    ("ihD4_top_s6", D4, None, S6, False),
    ("onecell_top_d4", S6, None, D4, True),
    ("onecell_bot_d4", S6, D4, None, True),
]

SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def build():
    body = []
    for i, (name, ih, r1b, r2t, one) in enumerate(ARMS):
        brk = '<w:pPr><w:pageBreakBefore/></w:pPr>' if i else ""
        body.append("<w:p>" + brk + txt(mark(i, "A")) + "</w:p>")
        body.append(tbl(ih, r1b, r2t, one))
        body.append("<w:p>" + txt(mark(i, "B")) + "</w:p>")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document %s><w:body>%s%s</w:body></w:document>' % (NS, "".join(body), SECT))
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          "</Types>")
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            "</Relationships>")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
    print("wrote", OUT)
    return OUT


def word_rows(path):
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
        pg = doc[pi]
        for blk in pg.get_text("dict")["blocks"]:
            if blk.get("type", 0) != 0:
                continue
            for ln in blk.get("lines", []):
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if t:
                    out.append((pi * 10000 + min(s["bbox"][1] for s in ln["spans"]),
                                "T", t, 0.0, 0.0, 0.0))
        for dd in pg.get_drawings():
            rc = dd["rect"]
            if rc.height < 8 and rc.width > 20:
                out.append((pi * 10000 + rc.y0, "B", "", rc.height, rc.x0, rc.x1))
    return out


def oxi_rows(path):
    exe = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
    tmp = Path(tempfile.mkdtemp())
    dump = tmp / "d.json"
    subprocess.run([str(exe), str(path), str(tmp / "p"), "110", "--dump-layout=%s" % dump],
                   check=True, capture_output=True)
    import json
    d = json.load(open(dump, encoding="utf-8"))
    out = []
    for pi, pg in enumerate(d["pages"]):
        for e in pg["elements"]:
            y = pi * 10000 + e.get("y", 0.0)
            if e.get("type") == "text" and (e.get("text") or "").strip():
                out.append((y, "T", (e.get("text") or "").strip(), 0.0, 0.0, 0.0))
            elif e.get("type") == "border" and (e.get("w") or 0) > 20:
                out.append((y, "B", "", 0.0, e.get("x", 0.0), e.get("x", 0.0) + (e.get("w") or 0)))
    return out


def summarize(rows, tag):
    rows.sort()
    print("--- %s ---" % tag)
    res = {}
    for i, (name, ih, r1b, r2t, one) in enumerate(ARMS):
        ya = [r[0] for r in rows if r[1] == "T" and mark(i, "A") in r[2]]
        yb = [r[0] for r in rows if r[1] == "T" and mark(i, "B") in r[2]]
        if not ya or not yb:
            print("  %-16s MISSING" % name)
            continue
        a, b = min(ya), min(yb)
        raw = sorted({(round(r[0], 2), round(r[3], 2), round(r[4], 1), round(r[5], 1))
                      for r in rows if r[1] == "B" and a < r[0] < b})
        # 1) merge rects that share a y (one per COLUMN) into a single line whose
        #    span is their union — otherwise every edge looks like a 0.0 pitch.
        lines = []
        for y, h, x0, x1 in raw:
            if lines and abs(y - lines[-1][0]) < 0.06:
                lines[-1][1] = max(lines[-1][1], h)
                lines[-1][2] = min(lines[-1][2], x0)
                lines[-1][3] = max(lines[-1][3], x1)
            else:
                lines.append([y, h, x0, x1])
        # 2) a double's two companions sit < 1.6pt apart -> one logical edge
        edges = []
        for y, h, x0, x1 in lines:
            if edges and y - edges[-1][-1][0] < 1.6:
                edges[-1].append([y, h, x0, x1])
            else:
                edges.append([[y, h, x0, x1]])
        pitches = [round(edges[k + 1][0][0] - edges[k][0][0], 2)
                   for k in range(len(edges) - 1)]
        tested = edges[EDGE_ROW + 1] if len(edges) > EDGE_ROW + 1 else None
        desc = "-"
        if tested:
            extent = max(y + h for y, h, _, _ in tested) - tested[0][0]
            span = max(x1 for _, _, _, x1 in tested) - min(x0 for _, _, x0, _ in tested)
            desc = "%dline extent=%.2f span=%.1f" % (len(tested), extent, span)
        print("  %-16s advance %7.2f  pitches %-28s edge: %s"
              % (name, b - a, str(pitches), desc))
        res[name] = (b - a, pitches, desc)
    return res


if __name__ == "__main__":
    p = build()
    w = summarize(word_rows(p), "WORD") if "--measure" in sys.argv else None
    o = summarize(oxi_rows(p), "OXI") if "--oxi" in sys.argv else None
    if w and o:
        print("--- DIFF (oxi - word) ---")
        for arm in ARMS:
            n = arm[0]
            if n in w and n in o:
                print("  %-16s d_advance %+7.2f  word_pitches %s  oxi_pitches %s"
                      % (n, o[n][0] - w[n][0], w[n][1], o[n][1]))
