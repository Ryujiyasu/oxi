"""Minimal repro: hanging-indent paragraph — where does Word place line 1 vs line 2+?

`<w:ind w:left="180" w:hanging="180"/>` = indent_left=180tw, first_line=-180tw (hanging).
  - Continuation lines sit at x = margin + 180tw = 9pt.
  - First line — question: does it sit at x = margin + 180tw + (-180tw) = 0tw (hanging)?
    Or at x = margin + 180tw = 9pt (same as continuation)?

Measure via Word COM. Create doc with two hanging-indent paragraphs (each long enough
to wrap 2 lines). Use COM Range.Information to read the x-coordinate of each line.
"""
import os, sys, time, zipfile
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CT = """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

RELS = """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

FONT_RPR = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/>'

# Two wide paragraphs that wrap onto 2+ lines, with hanging indent
LONG_TEXT_A = "第１条（目的）本契約は、委託業務に関する条件を定めることを目的とする。ただし、個別契約において別途定める場合はその限りではない。"
LONG_TEXT_B = "第２条（定義）本契約において使用する用語の定義は、契約書の各条項に示すところによる。"

def hanging_para(text, ind_left_tw, hang_tw):
    """Paragraph with w:ind left=X hanging=Y. text wraps across lines."""
    return f'''<w:p>
  <w:pPr>
    <w:ind w:left="{ind_left_tw}" w:hanging="{hang_tw}"/>
    <w:rPr>{FONT_RPR}</w:rPr>
  </w:pPr>
  <w:r><w:rPr>{FONT_RPR}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>
</w:p>'''

def normal_para(text):
    return f'<w:p><w:pPr><w:rPr>{FONT_RPR}</w:rPr></w:pPr><w:r><w:rPr>{FONT_RPR}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

body = f"""
{normal_para("Reference paragraph (no indent) for baseline x.")}
{hanging_para(LONG_TEXT_A, 180, 180)}
{hanging_para(LONG_TEXT_B, 180, 180)}
{hanging_para("短い行 (one line, hanging=180) — sits where?", 180, 180)}
"""

DOC_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body}
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>
</w:body>
</w:document>"""

DOCX = os.path.abspath("pipeline_data/hanging_indent_test.docx")
os.makedirs(os.path.dirname(DOCX), exist_ok=True)

with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/document.xml", DOC_XML)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
time.sleep(0.3)

# Sample every character's x, y — detect line breaks by y changes
# This is slow but precise. Limit to 300 chars.
print(f"Paragraph count: {doc.Paragraphs.Count}")
print("Margin expected: 851tw = 42.55pt from page left.")
print()
for p_idx in range(1, min(doc.Paragraphs.Count, 5) + 1):
    para = doc.Paragraphs(p_idx)
    rng = para.Range
    text = rng.Text.replace('\r', '¶')[:60]
    print(f'--- Para {p_idx}: "{text}" ---')
    # Get character-level coordinates via Duplicate + Move
    chars = rng.Characters
    n = min(chars.Count, 100)
    lines_seen = []  # list of (first_char_x, first_char_y)
    last_y = None
    for c in range(1, n + 1):
        ch = chars(c)
        x = ch.Information(5)   # wdHorizontalPositionRelativeToPage (pt)
        y = ch.Information(6)   # wdVerticalPositionRelativeToPage (pt)
        if last_y is None or abs(y - last_y) > 5:
            lines_seen.append((c, x, y, ch.Text))
            last_y = y
    for (cidx, lx, ly, txt) in lines_seen:
        print(f"    line-start char#{cidx} '{txt}' x={lx:.2f}pt y={ly:.2f}pt")

doc.Close(SaveChanges=False)
word.Quit()
