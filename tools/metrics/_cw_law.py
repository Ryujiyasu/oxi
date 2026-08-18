# -*- coding: utf-8 -*-
"""What is Word's line-wrap budget inside a table cell?

`wrap_base` in layout/mod.rs is an eight-condition allowlist: a cell matching
one of them wraps at cell_w - pad_l - pad_r, a cell matching none wraps at the
bare cell_w. An allowlist is the shape a rule takes when it was never derived,
and the documents that disagree about it (kojin wants the whole padding gone,
34140b/04b88e want half) carry no XML discriminator to separate them.

So measure the law. Each arm is a document of single-cell tables whose width is
swept in 0.5pt steps; every cell holds the same run of fullwidth CJK (one em per
character) left-aligned, so the count on the first line is floor(budget / em).
The width where that count steps k -> k+1 pins the budget to +-half a step:

    budget(w*) = (k+1) * em    =>    C(w) = w - budget = w* - (k+1)*em

Sweeping the cell margins and the border weight across arms says what C is made
of. The vertical rules are read out of the PDF drawings, so the column geometry
is measured too, not assumed.

    python _cw_law.py            # generate, export through Word, measure
    python _cw_law.py --keep     # reuse whatever is already exported
    python _cw_law.py --arms A,B # only some arms
"""
import os
import sys
import zipfile
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
OUT = os.path.join(REPO, "pipeline_data", "_cw_law")
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MARK = "甲"          # 甲 - first character of every cell, so a line that
FILL = "亜"          # 亜   starts with it is a cell start
NCHAR = 16
SZ = 21                  # half-points -> 10.5pt nominal em (Word draws 10.56)
# Two 1-twip windows rather than one coarse sweep: the advance and the constant
# trade off against each other inside a single window (a 0.5pt grid cannot tell
# 10.50pt/char + C=11.0 from 10.56pt/char + C=10.3), and only a long baseline
# between two far-apart transitions separates them.
# Each window is wider than one em so that it holds a transition whatever the
# constant turns out to be -- the arms move the transitions around.
WINDOWS = [(1000.0, 1250.0, 1.0), (2250.0, 2500.0, 1.0)]   # twips: 50..62.5pt, 112.5..125pt

# name -> (marL, marR, border sz eighths-of-a-point, tblLayout, jc, docGrid)
def arm(marl=108, marr=108, bd=4, layout="fixed", jc="left", grid="lines",
        compat="15"):
    return dict(marl=marl, marr=marr, bd=bd, layout=layout, jc=jc, grid=grid,
                compat=compat)


ARMS = {
    "A": arm(),
    "B": arm(marl=0, marr=0),
    "C": arm(marl=216, marr=216),
    "D": arm(marl=108, marr=432),
    "E": arm(bd=0),
    "F": arm(bd=24),
    "G": arm(marl=0, marr=0, bd=0),
    "H": arm(marl=0, marr=0, bd=24),
    "I": arm(jc="both"),
    "J": arm(jc="center"),
    "K": arm(jc="right"),
    "L": arm(layout="auto"),
    "M": arm(compat="11"),
    "N": arm(grid="linesAndChars"),
}
ARM_NOTE = {
    "A": "the corpus default (108/108, hairline rule)",
    "B": "no cell margin at all",
    "C": "double cell margin",
    "D": "asymmetric: does the RIGHT margin count the same as the left?",
    "E": "no borders: does the rule weight enter the budget?",
    "F": "3pt borders: same question, the other way",
    "G": "no margin AND no border: is the 0.24pt inset the rule's half-width?",
    "H": "no margin, 3pt border: same question, the other way",
    "I": "justified -- the alignment the allowlist singles out",
    "J": "centred -- the alignment the allowlist EXCLUDES",
    "K": "right-aligned -- likewise excluded",
    "L": "autofit: does the drawn column still equal the declared width?",
    "M": "legacy compatibilityMode 11, as most of the corpus is",
    "N": "a character grid as well as a line grid",
}


def widths():
    ws = []
    for w0, w1, step in WINDOWS:
        w = w0
        while w <= w1 + 1e-6:
            ws.append(w)
            w += step
    return ws


def build(out, marl, marr, bd, layout, jc, grid, compat):
    font = "ＭＳ 明朝"      # ＭＳ 明朝
    rpr = (f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
           f'<w:sz w:val="{SZ}"/>')
    if bd:
        border = ('<w:tblBorders>'
                  + ''.join(f'<w:{e} w:val="single" w:sz="{bd}" w:space="0" w:color="000000"/>'
                            for e in ("top", "left", "bottom", "right", "insideH", "insideV"))
                  + '</w:tblBorders>')
    else:
        border = ''
    tcmar = ('<w:tblCellMar>'
             '<w:top w:w="0" w:type="dxa"/>'
             f'<w:left w:w="{marl}" w:type="dxa"/>'
             '<w:bottom w:w="0" w:type="dxa"/>'
             f'<w:right w:w="{marr}" w:type="dxa"/>'
             '</w:tblCellMar>')
    text = MARK + FILL * (NCHAR - 1)
    tbls, n = [], 0
    for w in widths():
        cw = int(round(w))
        tbls.append(
            f'<w:tbl><w:tblPr><w:tblW w:w="{cw}" w:type="dxa"/>'
            f'<w:tblInd w:w="0" w:type="dxa"/>'
            f'<w:tblLayout w:type="{layout}"/>{border}{tcmar}</w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{cw}"/></w:tblGrid>'
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{cw}" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:pPr><w:jc w:val="{jc}"/><w:rPr>{rpr}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t>{text}</w:t></w:r>'
            f'</w:p></w:tc></w:tr></w:tbl>'
            f'<w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr></w:p>')
        n += 1
    docgrid = '' if grid == "none" else f'<w:docGrid w:type="{grid}" w:linePitch="360"/>'
    sectpr = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
              f'w:header="851" w:footer="992" w:gutter="0"/>{docgrid}</w:sectPr>')
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'<w:body>{"".join(tbls)}{sectpr}</w:body></w:document>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
              f'<w:sz w:val="{SZ}"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
              '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:compat><w:compatSetting w:name="compatibilityMode" '
                f'w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/>'
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
    return n


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


def read(pdf):
    """Per cell: the characters of every line, plus the vertical rules around it."""
    import fitz
    cells, cur = [], None
    for pg in fitz.open(pdf):
        rules = []                      # (x, y0, y1) vertical rules on this page
        for dr in pg.get_drawings():
            for it in dr["items"]:
                if it[0] == "l":
                    p, q = it[1], it[2]
                    if abs(p.x - q.x) < 0.3 and abs(p.y - q.y) > 1.0:
                        rules.append((p.x, min(p.y, q.y), max(p.y, q.y)))
                elif it[0] == "re":
                    r = it[1]
                    if r.width < 3.0 and r.height > 1.0:
                        rules.append((r.x0 + r.width / 2, r.y0, r.y1))
        lines = []
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                chars = []
                for sp in ln["spans"]:
                    for ch in sp["chars"]:
                        if ch["c"].strip():
                            chars.append((ch["c"], ch["bbox"][0], ch["bbox"][2]))
                if chars:
                    lines.append((round(ln["bbox"][1], 1), round(ln["bbox"][3], 1), chars))
        lines.sort(key=lambda t: (t[0], t[2][0][1]))
        for y0, y1, chars in lines:
            if chars[0][0] == MARK:
                cur = dict(lines=[], rules=None, y0=y0)
                cells.append(cur)
                near = sorted(x for x, ry0, ry1 in rules if ry0 - 2 <= y0 and ry1 >= y1 - 2)
                if len(near) >= 2:
                    cur["rules"] = (near[0], near[-1])
            if cur is not None:
                cur["lines"].append(chars)
    return cells


def advances(chars):
    return [chars[i + 1][1] - chars[i][1] for i in range(len(chars) - 1)]


def solve(obs, a_lo=10.30, a_hi=10.80, c_lo=-4.0, c_hi=30.0, grid=0.005):
    """Feasible (advance, constant) for `k chars fit iff k*a + C <= w`.

    Every cell contributes two inequalities, so the answer is a region, not a
    point; reporting its bounds is the honest form of the measurement."""
    ai = int(round((a_hi - a_lo) / grid))
    ci = int(round((c_hi - c_lo) / grid))
    feas = []
    for i in range(ai + 1):
        a = a_lo + i * grid
        # k*a + C <= w  and  (k+1)*a + C > w   =>   C in (w-(k+1)a, w-k*a]
        lo = max(w - (k + 1) * a for w, k in obs)
        hi = min(w - k * a for w, k in obs)
        if lo < hi - 1e-9 and hi >= c_lo and lo <= c_hi:
            feas.append((a, max(lo, c_lo), min(hi, c_hi)))
    return feas


def analyse(name, cells, marl, marr, bd):
    adv = sorted(a for c in cells for ln in c["lines"] for a in advances(ln))
    med = adv[len(adv) // 2] if adv else float("nan")
    ws = widths()
    print(f"\n=== arm {name}: marL={marl / 20:.2f}pt marR={marr / 20:.2f}pt "
          f"border={bd / 8:.2f}pt  ({ARM_NOTE[name]})")
    print(f"    cells read {len(cells)} / {len(ws)};  drawn advance median "
          f"{med:.3f}pt (nominal em {SZ / 2.0:.2f})")
    if len(cells) != len(ws):
        print("    !! cell count mismatch -- widths cannot be assigned, stopping")
        return None
    rows = [((w / 20.0), len(c["lines"][0]), c["lines"][0][0][1], c["rules"])
            for w, c in zip(ws, cells)]
    ins = sorted(x0 - r[0] for w, n, x0, r in rows if r)
    if ins:
        print(f"    left inset (text x0 - left rule): median {ins[len(ins) // 2]:.2f}pt "
              f"[{ins[0]:.2f}..{ins[-1]:.2f}]   (marL = {marl / 20:.2f})")
    dev = sorted(r[1] - r[0] - w for w, n, x0, r in rows if r)
    if dev:
        print(f"    rule span - nominal width: median {dev[len(dev) // 2]:+.2f}pt "
              f"[{dev[0]:+.2f}..{dev[-1]:+.2f}]")
    prev = None
    for w, n, x0, r in rows:
        if prev is not None and n != prev:
            print(f"      transition {prev}->{n} between {w - WINDOWS[0][2] / 20:.2f} "
                  f"and {w:.2f}pt")
        prev = n
    feas = solve([(w, n) for w, n, x0, r in rows])
    if not feas:
        print("    !! no (advance, constant) explains every cell -- the model is wrong")
        return None
    a_lo, a_hi = feas[0][0], feas[-1][0]
    c_lo = min(f[1] for f in feas)
    c_hi = max(f[2] for f in feas)
    mid = feas[len(feas) // 2]
    print(f"    => advance in [{a_lo:.3f}, {a_hi:.3f}]  constant in "
          f"({c_lo:.3f}, {c_hi:.3f}]")
    print(f"       at the drawn advance {med:.3f}: constant in "
          f"({max(w - (k + 1) * med for w, k in [(r[0], r[1]) for r in rows]):.3f}, "
          f"{min(w - k * med for w, k in [(r[0], r[1]) for r in rows]):.3f}]"
          f"   [marL+marR = {(marl + marr) / 20:.2f}, marL = {marl / 20:.2f}, "
          f"border = {bd / 8:.2f}]")
    return (a_lo, a_hi, c_lo, c_hi, mid, med)


RENDERER = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                        "oxi-gdi-renderer.exe")


def oxi_counts(docx):
    """First-line character count per cell, straight out of Oxi's own layout."""
    import json
    import subprocess
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        dump = os.path.join(td, "d.json")
        subprocess.run([RENDERER, docx, os.path.join(td, "p"), "96",
                        "--dump-layout=" + dump],
                       check=True, capture_output=True)
        d = json.load(open(dump, encoding="utf-8"))
    counts, started = [], False
    for pg in d["pages"]:
        for e in pg["elements"]:
            if e["type"] != "text" or not e.get("text"):
                continue
            if e["text"].startswith(MARK):
                counts.append(len(e["text"]))
                started = True
            elif not started:
                continue
    return counts


def compare(name, cfg, word_rows):
    ws = widths()
    docx = os.path.join(OUT, f"cw_{name}.docx")
    got = oxi_counts(docx)
    if len(got) != len(ws):
        print(f"    oxi: {len(got)} cells vs {len(ws)} widths -- cannot align")
        return
    feas = solve([(w / 20.0, k) for w, k in zip(ws, got)])
    agree = sum(1 for (w, n, x0, r), k in zip(word_rows, got) if n == k)
    # in twips, as Word does it -- in float points the exact fit falls an ulp short
    law = [(w / 20.0,
            int((w - max(cfg["marl"], cfg["bd"] * 1.25) - max(cfg["marr"], cfg["bd"] * 1.25))
                // (SZ * 10.0)))
           for w in ws]
    law_ok = sum(1 for (w, n, x0, r), (lw, lk) in zip(word_rows, law) if n == lk)
    print(f"    oxi agrees with Word on {agree}/{len(ws)} cells; "
          f"the derived law agrees on {law_ok}/{len(ws)}")
    if feas:
        print(f"    oxi's own law: advance in [{feas[0][0]:.3f},{feas[-1][0]:.3f}] "
              f"constant in ({min(f[1] for f in feas):.3f}, "
              f"{max(f[2] for f in feas):.3f}]")
    else:
        over = Counter(k - n for (w, n, x0, r), k in zip(word_rows, got))
        print("    oxi fits no single (advance, constant); chars over Word: "
              f"{dict(sorted(over.items()))}")


def main():
    os.makedirs(OUT, exist_ok=True)
    a = sys.argv
    want = (a[a.index("--arms") + 1].split(",") if "--arms" in a else list(ARMS))
    keep = "--keep" in a
    with_oxi = "--oxi" in a
    summary = {}
    for name in want:
        cfg = ARMS[name]
        docx = os.path.join(OUT, f"cw_{name}.docx")
        if not (keep and os.path.exists(docx)):
            n = build(docx, **cfg)
            print(f"built {docx}: {n} cells")
        cells = read(export(docx))
        summary[name] = analyse(name, cells, cfg["marl"], cfg["marr"], cfg["bd"])
        if with_oxi and summary[name]:
            ws = widths()
            compare(name, cfg,
                    [((w / 20.0), len(c["lines"][0]), c["lines"][0][0][1], c["rules"])
                     for w, c in zip(ws, cells)])
    print("\n=== budget = cell width - constant, by arm ===")
    print(f"  {'arm':<4}{'marL':>7}{'marR':>7}{'border':>8}"
          f"{'advance':>18}{'constant':>20}")
    for name, r in summary.items():
        if not r:
            continue
        cfg = ARMS[name]
        a_lo, a_hi, c_lo, c_hi, mid, med = r
        print(f"  {name:<4}{cfg['marl'] / 20:7.2f}{cfg['marr'] / 20:7.2f}"
              f"{cfg['bd'] / 8:8.2f}   [{a_lo:.3f},{a_hi:.3f}]   ({c_lo:7.3f},{c_hi:7.3f}]")


if __name__ == "__main__":
    main()
