"""Test both Latin and CJK hanging-indent to diagnose v1 result (all lines at same x)."""
import os, sys, time, zipfile
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CT = """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"""
RELS = """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""

TESTS = [
    # (label, ind_attrs, rpr, text_wraps_2_lines)
    ("Latin hanging=720tw (36pt)",
     'w:left="720" w:hanging="720"',
     '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>',
     "This is a long test paragraph that should wrap to multiple lines to see where the continuation line starts after the hanging indent takes effect. Keep going.",
    ),
    ("CJK MS Mincho hanging=180tw (9pt)",
     'w:left="180" w:hanging="180"',
     '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/>',
     "第１条（目的）本契約は、委託業務に関する条件を定めることを目的とする。ただし、個別契約において別途定める場合はその限りではない。"
    ),
    ("CJK firstLine=-180 (same as hanging=180)",
     'w:left="180" w:firstLine="-180"',
     '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/>',
     "第１条（目的）本契約は、委託業務に関する条件を定めることを目的とする。ただし、個別契約において別途定める場合はその限りではない。"
    ),
    ("CJK hanging=420tw (21pt = 2 chars)",
     'w:left="420" w:hanging="420"',
     '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/>',
     "第１条（目的）本契約は、委託業務に関する条件を定めることを目的とする。ただし、個別契約において別途定める場合はその限りではない。"
    ),
]

def make_para(ind_attrs, rpr, text):
    return f'<w:p><w:pPr><w:ind {ind_attrs}/><w:rPr>{rpr}</w:rPr></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

body = "\n".join(make_para(a, r, t) for _, a, r, t in TESTS)
DOC_XML = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>'

DOCX = os.path.abspath("pipeline_data/hanging_indent_v2.docx")
with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/document.xml", DOC_XML)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
time.sleep(0.3)

for i, (label, attrs, rpr, text) in enumerate(TESTS, 1):
    print(f"\n=== Test {i}: {label} ===")
    print(f"    XML attrs: {attrs}")
    para = doc.Paragraphs(i)
    chars = para.Range.Characters
    n = min(chars.Count, 120)
    last_y = None
    for c in range(1, n + 1):
        ch = chars(c)
        x = ch.Information(7)   # TextBoundary-relative (true glyph offset from margin)
        y = ch.Information(6)
        if last_y is None or abs(y - last_y) > 5:
            print(f"    line-start char#{c} '{ch.Text}' x={x:.2f}pt y={y:.2f}pt")
            last_y = y

doc.Close(SaveChanges=False)
word.Quit()
