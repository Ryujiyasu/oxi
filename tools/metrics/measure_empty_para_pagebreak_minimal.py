"""Minimal repro: empty paragraph with <w:br type="page"/> — does Word give it vertical space?

Three variants:
  A) Normal page break via paragraph with only <w:br type="page"/> run (0e7a-style)
  B) page_break_before on the next paragraph (cleaner OOXML pattern)
  C) no break at all (control)

For each, render via Word COM and read Paragraphs(n).Range.Information(6) (y on page)
and Information(3) (page number). Compare resulting Y layout.
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

def doc_xml(body_paras):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body_paras}
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>
</w:body>
</w:document>"""

FONT_RPR = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/>'

def text_para(text, page_break_before=False):
    pbb = '<w:pageBreakBefore/>' if page_break_before else ''
    return f'<w:p><w:pPr>{pbb}<w:rPr>{FONT_RPR}</w:rPr></w:pPr><w:r><w:rPr>{FONT_RPR}</w:rPr><w:t>{text}</w:t></w:r></w:p>'

def empty_para_with_br():
    """0e7a-style: paragraph with only <w:br type="page"/> run."""
    return f'<w:p><w:pPr><w:rPr>{FONT_RPR}</w:rPr></w:pPr><w:r><w:rPr>{FONT_RPR}</w:rPr><w:br w:type="page"/></w:r></w:p>'

VARIANTS = {
    "A_0e7a_style_inline_br": [
        text_para("Para1 on p1"),
        text_para("Para2 on p1"),
        empty_para_with_br(),   # empty para with inline br
        text_para("Para3 after break (should be p2)"),
        text_para("Para4 on p2"),
    ],
    "B_page_break_before": [
        text_para("Para1 on p1"),
        text_para("Para2 on p1"),
        text_para("Para3 after break (should be p2)", page_break_before=True),
        text_para("Para4 on p2"),
    ],
    "C_no_break": [
        text_para("Para1 on p1"),
        text_para("Para2 on p1"),
        text_para("Para3 on p1"),
        text_para("Para4 on p1"),
    ],
}

OUT_DIR = os.path.abspath("pipeline_data")
os.makedirs(OUT_DIR, exist_ok=True)

def build_docx(path, paras):
    body = "\n".join(paras)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc_xml(body))

def measure(path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    doc = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.3)
    rows = []
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        rng = p.Range
        page = rng.Information(3)
        y = rng.Information(6)
        txt = rng.Text.replace("\r", "").replace("\x0C", "[PAGEBREAK]")[:40]
        rows.append((i, page, y, txt))
    doc.Close(SaveChanges=False)
    word.Quit()
    return rows

for name, paras in VARIANTS.items():
    path = os.path.join(OUT_DIR, f"empty_para_repro_{name}.docx")
    build_docx(path, paras)
    rows = measure(path)
    print(f"\n=== {name} ===")
    prev_y = None
    prev_page = None
    for i, page, y, txt in rows:
        gap = ""
        if prev_y is not None and prev_page == page:
            gap = f"gap={y - prev_y:+.2f}"
        print(f"  P{i} page={page} y={y:7.2f} {gap:>12}  | {txt}")
        prev_y = y
        prev_page = page
