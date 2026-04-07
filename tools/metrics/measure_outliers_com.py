"""COM-measure exact line/para metrics for the 5 outlier docs."""
import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

DOCS = [
    "nested_bullet_08",
    "page_break_paragraph_spacing",
    "style_inheritance_complex_19",
    "image_text_wrap_complex_01",
    "mixed_font_line_height",
]

for d in DOCS:
    p = os.path.abspath(f"pipeline_data/docx/{d}.docx")
    doc = word.Documents.Open(p, ReadOnly=True)
    print(f"\n=== {d} ===")
    for i in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs(i)
        fmt = para.Format
        rng = para.Range
        # First-char Y
        try:
            y = rng.Characters(1).Information(6)  # wdVerticalPositionRelativeToPage
        except: y = -1
        sb = fmt.SpaceBefore
        sa = fmt.SpaceAfter
        ls = fmt.LineSpacing
        lsr = fmt.LineSpacingRule  # 0=Single, 1=1.5, 2=Double, 3=AtLeast, 4=Exactly, 5=Multiple
        font = rng.Characters(1).Font
        fname = font.NameFarEast or font.Name
        fsz = font.Size
        text = (rng.Text or "").replace("\r","").replace("\n","")[:30]
        print(f"  P{i:2d} y={y:7.2f} sb={sb:5.1f} sa={sa:5.1f} ls={ls:5.1f} lsr={lsr} font={fname[:14]:14s} sz={fsz}  {text!r}")
    doc.Close(False)

word.Quit()
