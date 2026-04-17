"""Measure Word's x-position per char for MS Gothic 12pt in a d77a-like context.
Check if ASCII digits '0'-'9' and fullwidth chars match Oxi's gdi_width_overrides.
"""
import os, sys, time, zipfile
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CT = """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"""
RELS = """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""

RPR = '<w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="24"/>'

# Test strings — include actual chars from d77a PARA 21 L1
tests = [
    ("digit_run", "0123456789"),
    ("d77a_L1_head", "平成26年6月19日に決定した"),
    ("fullwidth_latin", "ＷｅｂＷｅｂ"),
    ("common_cjk", "公共規範各省庁"),
]

body = "\n".join(f'<w:p><w:pPr><w:rPr>{RPR}</w:rPr></w:pPr><w:r><w:rPr>{RPR}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r></w:p>' for _, t in tests)
DOC = f'<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992"/></w:sectPr></w:body></w:document>'
path = os.path.abspath("pipeline_data/ms_gothic_12pt_widths.docx")
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/document.xml", DOC)

w = win32com.client.Dispatch("Word.Application"); w.Visible = False
doc = w.Documents.Open(path, ReadOnly=True); time.sleep(0.3)

for i, (label, text) in enumerate(tests, 1):
    print(f"\n=== {label}: {text!r} ===")
    chars = doc.Paragraphs(i).Range.Characters
    n = min(chars.Count, len(text))
    prev_x = None
    for c in range(1, n + 1):
        ch = chars(c)
        x = ch.Information(7)  # TextBoundary-relative
        txt = ch.Text
        cp = ' '.join(f'{ord(t):04X}' for t in txt)
        if prev_x is None:
            print(f"  c{c} '{txt}' (U+{cp}) x={x:7.2f}pt  (first)")
        else:
            adv = x - prev_x
            print(f"  c{c} '{txt}' (U+{cp}) x={x:7.2f}pt  advance_from_prev={adv:.2f}pt")
        prev_x = x

doc.Close(False); w.Quit()
