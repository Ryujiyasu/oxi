"""Build minimal docx fixtures for fn reserve algorithm measurement.

Constructs OOXML directly (bypass python-docx footnote limitation).

Scenarios:
  RA: many body lines + final para carrying 5 fn refs
      → tests widow/orphan + reserve interaction
  RB: same as RA but final para has 1 fn ref (small reserve)
      → comparison baseline for how reserve scales
  RC: body lines + 3 paras each with 1 fn (streaming)
      → tests if Word streams refs per-para
  RD: body lines + final para has 10 fn refs (extreme reserve demand)
      → forces widow/orphan
"""
import os, zipfile, io

OUT_DIR = r"tools\metrics\fn_reserve_repro"
os.makedirs(OUT_DIR, exist_ok=True)

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "cp": "http://schemas.openxmlformats.org/package/2006/content-types",
}

CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"""

ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"""

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="ＭＳ 明朝" w:cs="Times New Roman"/>
<w:sz w:val="21"/><w:szCs w:val="21"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr>
<w:spacing w:line="276" w:lineRule="auto"/>
</w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="FootnoteText">
<w:name w:val="footnote text"/>
<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
</w:style>
</w:styles>"""


def footnotes_xml(fn_texts):
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        # Standard separator footnotes
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    ]
    for i, t in enumerate(fn_texts, start=1):
        parts.append(
            f'<w:footnote w:id="{i}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
            f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
            f'<w:r><w:t xml:space="preserve"> {t}</w:t></w:r></w:p></w:footnote>'
        )
    parts.append("</w:footnotes>")
    return "".join(parts)


def para(text, fn_ids=None, page_break=False):
    parts = ["<w:p>"]
    if page_break:
        parts.append('<w:r><w:br w:type="page"/></w:r>')
    parts.append(f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>')
    if fn_ids:
        for fid in fn_ids:
            parts.append(
                f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                f'<w:footnoteReference w:id="{fid}"/></w:r>'
            )
    parts.append("</w:p>")
    return "".join(parts)


def document_xml(body_paras):
    body_inner = "".join(body_paras)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{body_inner}'
        # Section: A4 (11906 x 16838 twips), margins 1440 twips (~1 inch)
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:docGrid w:type="lines" w:linePitch="360"/>'
        '</w:sectPr>'
        '</w:body></w:document>'
    )


def build_docx(path, body_paras, fn_texts):
    """Assemble a minimal docx."""
    files = {
        "[Content_Types].xml": CONTENT_TYPES_XML,
        "_rels/.rels": ROOT_RELS_XML,
        "word/_rels/document.xml.rels": DOC_RELS_XML,
        "word/styles.xml": STYLES_XML,
        "word/settings.xml": SETTINGS_XML,
        "word/footnotes.xml": footnotes_xml(fn_texts),
        "word/document.xml": document_xml(body_paras),
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content)
    print(f"Saved {path}")


def filler_line(n):
    return (
        "これはフィラーの段落です。十分な長さがあるので"
        "改行が発生することを期待します。" * 3
    )


def scenario_RA():
    """10 body paras + final para with 5 fn refs."""
    paras = [para(filler_line(i)) for i in range(10)]
    paras.append(para("ラスト段落：", fn_ids=[1, 2, 3, 4, 5]))
    fn_texts = [f"Footnote {i+1} 本文" for i in range(5)]
    build_docx(os.path.join(OUT_DIR, "RA_10body_5fns_end.docx"), paras, fn_texts)


def scenario_RB():
    """10 body paras + final para with 1 fn ref."""
    paras = [para(filler_line(i)) for i in range(10)]
    paras.append(para("ラスト段落：", fn_ids=[1]))
    fn_texts = ["Footnote 1 本文"]
    build_docx(os.path.join(OUT_DIR, "RB_10body_1fn_end.docx"), paras, fn_texts)


def scenario_RC():
    """10 body paras + 3 paras each with 1 fn (streaming test)."""
    paras = [para(filler_line(i)) for i in range(10)]
    paras.append(para("段落 A：", fn_ids=[1]))
    paras.append(para("段落 B：", fn_ids=[2]))
    paras.append(para("段落 C：", fn_ids=[3]))
    fn_texts = [f"Footnote {i+1} 本文" for i in range(3)]
    build_docx(os.path.join(OUT_DIR, "RC_10body_3paras_1fn_each.docx"), paras, fn_texts)


def scenario_RD():
    """10 body paras + final para with 10 fn refs (extreme reserve)."""
    paras = [para(filler_line(i)) for i in range(10)]
    paras.append(para("ラスト段落（重いfn）：", fn_ids=list(range(1, 11))))
    fn_texts = [f"Footnote {i+1} 本文" for i in range(10)]
    build_docx(os.path.join(OUT_DIR, "RD_10body_10fns_end.docx"), paras, fn_texts)


if __name__ == "__main__":
    scenario_RA()
    scenario_RB()
    scenario_RC()
    scenario_RD()
    print("Done.")
