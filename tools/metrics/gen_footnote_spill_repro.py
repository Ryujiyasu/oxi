"""Scratch minimal docx testing whether Word splits long footnote bodies.

V1: Single very long footnote body (exceeds body area bottom)
V2: Many short footnotes totalling more space than body area allows
V3: Single footnote with 20+ lines of content

If Word splits ANY of these across pages, oxi-2's hypothesis stands.
If not, spill model is false — bug is elsewhere.
"""
import os, subprocess, sys, time, zipfile
from pathlib import Path

OUT_DIR = Path("pipeline_data/_footnote_spill_variants")
OUT_DIR.mkdir(parents=True, exist_ok=True)

CT = '''<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>'''

RELS = '''<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>'''

SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'


def make_docx(path, body_paras, footnotes):
    """Build a docx from scratch with given body paras and footnote definitions.
    footnotes: list of (fn_id, body_text) tuples.
    """
    doc_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
    for i, para in enumerate(body_paras):
        doc_xml += para
    doc_xml += SECT + '</w:body></w:document>'

    # Footnotes.xml
    fn_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    # Required separator/continuation entries
    fn_xml += '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
    fn_xml += '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    for fn_id, body_text in footnotes:
        fn_xml += f'<w:footnote w:id="{fn_id}"><w:p><w:pPr><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="18"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="18"/></w:rPr><w:footnoteRef/></w:r><w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="18"/></w:rPr><w:t>{body_text}</w:t></w:r></w:p></w:footnote>'
    fn_xml += '</w:footnotes>'

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/footnotes.xml", fn_xml)


def body_para(text, fn_ref_id=None):
    rfonts = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/>'
    ref = f'<w:r><w:rPr>{rfonts}<w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="{fn_ref_id}"/></w:r>' if fn_ref_id else ''
    return f'<w:p><w:pPr><w:rPr>{rfonts}</w:rPr></w:pPr><w:r><w:rPr>{rfonts}</w:rPr><w:t>{text}</w:t></w:r>{ref}</w:p>'


# V1: single very long footnote
long_text = 'あ' * 500  # ~500 chars @10pt fn font would span many lines
v1_body = [body_para(f'本文段落{i}。', fn_ref_id=1 if i == 0 else None) for i in range(3)]
v1_footnotes = [(1, long_text)]

# V2: many short footnotes at one reference page
v2_body = [body_para(f'本文{i}', fn_ref_id=i+1) for i in range(15)]
v2_footnotes = [(i+1, f'短脚注{i} ' + 'あ' * 50) for i in range(15)]

# V3: single footnote with 20 explicit lines
v3_lines = ['行' + str(i) + 'あ' * 30 for i in range(20)]
v3_footnotes_text = '。'.join(v3_lines)
v3_body = [body_para('テスト本文', fn_ref_id=1)]
v3_footnotes = [(1, v3_footnotes_text)]


def main():
    make_docx(OUT_DIR / "V1_single_long_fn.docx", v1_body, v1_footnotes)
    make_docx(OUT_DIR / "V2_many_short_fn.docx", v2_body, v2_footnotes)
    make_docx(OUT_DIR / "V3_one_fn_20lines.docx", v3_body, v3_footnotes)
    print(f"Generated 3 variants in {OUT_DIR}")


if __name__ == "__main__":
    main()
