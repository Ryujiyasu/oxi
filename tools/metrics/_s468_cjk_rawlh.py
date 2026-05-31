"""S468: directly measure CJK per-font raw line height, Word vs Oxi.

S467 INFERRED (from VSNAP gate regressions) that Oxi's CJK raw line heights
!= Word's, blocking the 0.75pt grid-snap model. This MEASURES it directly:

Two questions:
  Q1: Is Word's CJK line TOP on the 0.75pt (15-twip) grid, like Latin?
      (If yes -> grid model is universal, only Oxi's raw advance is wrong.
       If no  -> grid model is Latin-only.)
  Q2: Does Oxi's CJK per-line advance match Word's?  By how much per line?

Pure CJK body, single spacing, sa0 (no spacing confound), NO docGrid.
Cross several fonts/sizes. Word: collapsed-start Information(6) (R30).
Oxi: GDI --dump-layout per-element y.
"""
import os, io, json, subprocess
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
import win32com.client as win32

REPO = r"C:\Users\ryuji\oxi-main"
OUT = os.path.join(REPO, "tools", "golden-test", "repros", "cjk_rawlh")
os.makedirs(OUT, exist_ok=True)
RENDERER = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")
TMP = r"C:\Users\ryuji\AppData\Local\Temp"
VPOS = 6
PAGE = 3
CJK = "日本語のテキスト本文確認用の行"  # 日本語のテキスト本文確認用の行


def set_east_asian(style, font):
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = rpr.makeelement(qn('w:rFonts'), {})
        rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), font)
    rfonts.set(qn('w:ascii'), font)
    rfonts.set(qn('w:hAnsi'), font)


def set_docgrid(sec, line_pitch, char_space=0, gtype="linesAndChars"):
    sectPr = sec._sectPr
    dg = sectPr.find(qn('w:docGrid'))
    if dg is None:
        dg = sectPr.makeelement(qn('w:docGrid'), {})
        sectPr.append(dg)
    dg.set(qn('w:type'), gtype)
    dg.set(qn('w:linePitch'), str(line_pitch))
    if char_space:
        dg.set(qn('w:charSpace'), str(char_space))


def make(fname, font, size, mult=1.0, after=0, n=12, docgrid=None):
    d = docx.Document()
    sec = d.sections[0]
    sec.page_width = Inches(8.5); sec.page_height = Inches(11)
    sec.top_margin = Inches(1); sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1.25); sec.right_margin = Inches(1.25)
    if docgrid is not None:
        set_docgrid(sec, *docgrid)
    st = d.styles["Normal"]
    st.font.size = Pt(size)
    set_east_asian(st, font)
    pf = st.paragraph_format
    if mult == 1.0:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    else:
        pf.line_spacing = mult; pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.space_after = Pt(after); pf.space_before = Pt(0)
    for i in range(n):
        d.add_paragraph(CJK)
    p = os.path.join(OUT, fname)
    d.save(p)
    return p


def grid_res(y, pitch=0.75):
    return round(abs(round(y / pitch) * pitch - y), 3)


def word_ys(word, path):
    doc = word.Documents.Open(path, ReadOnly=True)
    ys = []
    for para in doc.Paragraphs:
        rng = para.Range
        st = doc.Range(rng.Start, rng.Start)
        if st.Information(PAGE) != 1:
            continue
        if not para.Range.Text.strip():
            continue
        ys.append(round(st.Information(VPOS), 3))
    doc.Close(False)
    return ys


def oxi_ys(path):
    dump = os.path.join(TMP, "s468_dump.json")
    out_prefix = os.path.join(TMP, "s468_out")
    subprocess.run([RENDERER, path, out_prefix, "150", "--dump-layout=" + dump],
                   capture_output=True, text=True)
    d = json.load(io.open(dump, encoding="utf-8"))
    seen = set()
    for pg in d["pages"]:
        if pg["page"] != 1:
            continue
        for el in pg["elements"]:
            if el.get("type") != "text":
                continue
            if not (el.get("text") or "").strip():
                continue
            seen.add(round(el["y"], 2))  # dedupe per-glyph els -> unique line tops
    return sorted(seen)


def advances(ys):
    return [round(ys[i] - ys[i - 1], 3) for i in range(1, len(ys))]


def main():
    variants = [
        dict(fname="M_105_single.docx", font="MS Mincho", size=10.5),
        dict(fname="M_9_single.docx", font="MS Mincho", size=9),
        dict(fname="M_12_single.docx", font="MS Mincho", size=12),
        dict(fname="G_105_single.docx", font="MS Gothic", size=10.5),
        dict(fname="Meiryo_105_single.docx", font="Meiryo", size=10.5),
        dict(fname="M_105_m115.docx", font="MS Mincho", size=10.5, mult=1.15),
        dict(fname="M_105_sa10.docx", font="MS Mincho", size=10.5, after=10),
        # docGrid linesAndChars (dominant CJK regime) -- line height becomes grid pitch
        dict(fname="M_105_dg360.docx", font="MS Mincho", size=10.5, docgrid=(360, 0)),
        dict(fname="M_105_dg360_cs.docx", font="MS Mincho", size=10.5, docgrid=(360, 1453)),
        dict(fname="M_105_dg312.docx", font="MS Mincho", size=10.5, docgrid=(312, 0)),
        dict(fname="G_9_dg312.docx", font="MS Gothic", size=9, docgrid=(312, 0)),
    ]
    paths = []
    for kw in variants:
        fn = kw["fname"]
        font = kw["font"]; size = kw["size"]
        mult = kw.get("mult", 1.0); after = kw.get("after", 0)
        paths.append((make(fn, font, size, mult, after, docgrid=kw.get("docgrid")),
                      fn, font, size, mult, after))

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    out = []
    for p, fn, font, size, mult, after in paths:
        wy = word_ys(word, p)
        oy = oxi_ys(p)
        wa = advances(wy)
        oa = advances(oy)
        wa_mean = round(sum(wa) / len(wa), 3) if wa else 0
        oa_mean = round(sum(oa) / len(oa), 3) if oa else 0
        # grid residual of Word tops (the Q1 answer)
        wres = [grid_res(y) for y in wy]
        wres_max = max(wres) if wres else 0
        ores = [grid_res(y) for y in oy]
        n = min(len(wy), len(oy))
        out.append("=== %s  (%s %.1fpt mult=%.2f sa=%d) ===" % (fn, font, size, mult, after))
        out.append("  Word advances: %s  mean=%.3f" % (wa[:6], wa_mean))
        out.append("  Oxi  advances: %s  mean=%.3f" % (oa[:6], oa_mean))
        out.append("  per-line drift (oxi_adv - word_adv) mean = %+.3f  => over %d lines = %+.2f"
                   % (oa_mean - wa_mean, n, (oa_mean - wa_mean) * n))
        out.append("  Word top 0.75-grid residual: max=%.3f  %s" %
                   (wres_max, "ON-GRID" if wres_max < 0.02 else "OFF-GRID"))
        out.append("  Oxi  top 0.75-grid residual: max=%.3f" % (max(ores) if ores else 0))
        out.append("  Word ys[:5]=%s" % [round(y, 2) for y in wy[:5]])
        out.append("  Oxi  ys[:5]=%s" % [round(y, 2) for y in oy[:5]])
        out.append("")
    word.Quit()
    txt = "\n".join(out)
    io.open(os.path.join(REPO, "tools", "metrics", "_s468_cjk_rawlh.out"), "w", encoding="utf-8").write(txt)
    print(txt.encode("ascii", "replace").decode())


if __name__ == "__main__":
    main()
