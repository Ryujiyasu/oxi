"""Measure Word's empty paragraph heights at various font sizes.

2ea81 p2 has idx=120 (fs=14pt empty) and idx=121 (fs=16pt empty) that
Oxi may be rendering too short, causing page break to miss the fs=16pt
heading at idx=122.

Test: emit paragraphs with just pPr/rPr font size, no text, measure gap.
"""
import os, sys, time, zipfile
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CT = """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"""
RELS = """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""

def mincho(sz_half):
    return f'<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="{sz_half}"/>'

# Test each font size: build an empty para with pPr/rPr + sz. Measure gap to next para.
sizes = [(21, '10.5pt'), (24, '12pt'), (28, '14pt'), (32, '16pt'), (36, '18pt'), (40, '20pt')]

paras = []
# Anchor paragraph with normal text
paras.append(f'<w:p><w:pPr><w:rPr>{mincho(21)}</w:rPr></w:pPr><w:r><w:rPr>{mincho(21)}</w:rPr><w:t>anchor</w:t></w:r></w:p>')
# For each size, emit an empty paragraph with that font + normal follow-up
for (sz, label) in sizes:
    paras.append(f'<w:p><w:pPr><w:rPr>{mincho(sz)}</w:rPr></w:pPr></w:p>')
    paras.append(f'<w:p><w:pPr><w:rPr>{mincho(21)}</w:rPr></w:pPr><w:r><w:rPr>{mincho(21)}</w:rPr><w:t>after-{label}</w:t></w:r></w:p>')

body = '\n'.join(paras)
DOC = f'<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992"/></w:sectPr></w:body></w:document>'
path = os.path.abspath("pipeline_data/empty_para_heights.docx")
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/document.xml", DOC)

w = win32com.client.Dispatch("Word.Application"); w.Visible = False
doc = w.Documents.Open(path, ReadOnly=True); time.sleep(0.4)

print(f"Paras: {doc.Paragraphs.Count}")
print(f"{'#':>3} {'y':>7} {'gap':>6} (Word's empty paragraph rendering)")
prev_y = None
for i in range(1, doc.Paragraphs.Count + 1):
    p = doc.Paragraphs(i).Range
    y = p.Information(6)
    gap = "-" if prev_y is None else f"{y - prev_y:.2f}"
    # Label
    text = p.Text.replace('\r', '').strip()[:25]
    if not text: text = "(empty)"
    print(f"{i:>3} {y:>7.2f} {gap:>6}   {text}")
    prev_y = y

doc.Close(False); w.Quit()
