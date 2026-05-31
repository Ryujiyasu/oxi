"""S467 minimal repro: isolate the vertical grid-snap behavior.
Generate controlled body-only docx variants (no tables/lists) and measure
each paragraph's exact Y via Word COM (collapsed-start Information(6), R30).
Goal: confirm Word snaps line tops to the 0.75pt (15-twip = 96-DPI pixel)
grid, determine whether it snaps cumulative position vs advance, and the
rounding rule."""
import os, io
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_LINE_SPACING
import win32com.client as win32

OUT = r"C:\Users\ryuji\oxi-main\tools\golden-test\repros\grid_snap"
os.makedirs(OUT, exist_ok=True)
VPOS = 6
PAGE = 3


def make(fname, font="Calibri", size=11, mult=1.15, after=10, ndocgrid=True, n=14):
    d = docx.Document()
    sec = d.sections[0]
    sec.page_width = Inches(8.5); sec.page_height = Inches(11)
    sec.top_margin = Inches(1); sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1.25); sec.right_margin = Inches(1.25)
    st = d.styles["Normal"]
    st.font.name = font; st.font.size = Pt(size)
    pf = st.paragraph_format
    pf.line_spacing = mult; pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.space_after = Pt(after); pf.space_before = Pt(0)
    for i in range(n):
        d.add_paragraph("Line %02d the quick brown fox jumps" % i)
    p = os.path.join(OUT, fname)
    d.save(p)
    return p


def measure(word, path, tag):
    doc = word.Documents.Open(path, ReadOnly=True)
    ys = []
    for para in doc.Paragraphs:
        rng = para.Range
        st = doc.Range(rng.Start, rng.Start)
        if st.Information(PAGE) != 1:
            continue
        t = para.Range.Text.strip()
        if not t:
            continue
        ys.append((round(st.Information(VPOS), 3), t[:14]))
    doc.Close(False)
    lines = ["=== %s ===" % tag]
    prev = None
    for y, t in ys:
        gap = round(y - prev, 3) if prev is not None else 0.0
        on075 = "grid" if abs(round(y / 0.75) * 0.75 - y) < 0.005 else "OFF"
        lines.append("y=%8.3f  gap=%+7.3f  %s  %s" % (y, gap, on075, t))
        prev = y
    return "\n".join(lines)


def main():
    variants = [
        ("body_cal11_m115_sa10.docx", dict(font="Calibri", size=11, mult=1.15, after=10)),
        ("body_cal11_m115_sa0.docx", dict(font="Calibri", size=11, mult=1.15, after=0)),
        ("body_cal11_single_sa10.docx", dict(font="Calibri", size=11, mult=1.0, after=10)),
        ("body_times12_m1_sa6.docx", dict(font="Times New Roman", size=12, mult=1.0, after=6)),
        ("body_cal11_m2_sa0.docx", dict(font="Calibri", size=11, mult=2.0, after=0)),
    ]
    paths = [(make(fn, **kw), fn) for fn, kw in variants]
    word = win32.gencache.EnsureDispatch("Word.Application"); word.Visible = False
    out = []
    for p, fn in paths:
        out.append(measure(word, p, fn))
    word.Quit()
    txt = "\n\n".join(out)
    io.open(r"C:\Users\ryuji\oxi-main\tools\metrics\_s467_repro_grid.out", "w", encoding="utf-8").write(txt)
    print(txt.encode("ascii", "replace").decode())


if __name__ == "__main__":
    main()
