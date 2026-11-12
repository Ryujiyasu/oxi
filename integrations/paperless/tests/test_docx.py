# SPDX-License-Identifier: MIT OR Apache-2.0
"""Text and properties read from .docx without Paperless: pytest tests/."""

import io
import zipfile
from pathlib import Path

from paperless_oxidocs.docx import core_properties, created_date, docx_text

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'


def make_docx(tmp_path: Path, body: str, parts: dict[str, str] | None = None, core: str | None = None) -> Path:
    path = tmp_path / "t.docx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", f"<w:document {NS}><w:body>{body}</w:body></w:document>")
        for name, xml in (parts or {}).items():
            z.writestr(name, xml)
        if core is not None:
            z.writestr("docProps/core.xml", core)
    return path


def p(text: str) -> str:
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


def test_paragraphs_tables_and_text_boxes_read_once(tmp_path):
    box = (
        "<w:p><w:r><mc:AlternateContent><mc:Choice><w:txbxContent>" + p("In the box") + "</w:txbxContent></mc:Choice>"
        "<mc:Fallback><w:txbxContent>" + p("In the box") + "</w:txbxContent></mc:Fallback></mc:AlternateContent></w:r>"
        "<w:r><w:t>Anchor</w:t><w:tab/><w:t>after tab</w:t></w:r></w:p>"
    )
    table = "<w:tbl><w:tr><w:tc>" + p("Yes") + "</w:tc><w:tc>" + p("Yes") + "</w:tc></w:tr></w:tbl>"
    path = make_docx(tmp_path, p("Title") + box + table + "<w:p><w:del><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>")
    assert docx_text(path).split("\n") == ["Title", "Anchor\tafter tab", "In the box", "Yes", "Yes"]


def test_headers_repeat_once_and_follow_the_body(tmp_path):
    hdr = f"<w:hdr {NS}>" + p("Company Ltd") + "</w:hdr>"
    path = make_docx(tmp_path, p("Body"), {"word/header1.xml": hdr, "word/header2.xml": hdr, "word/footnotes.xml": f"<w:footnotes {NS}><w:footnote>" + p("A note") + "</w:footnote></w:footnotes>"})
    assert docx_text(path).split("\n") == ["Body", "Company Ltd", "A note"]


def test_core_properties_and_created_date(tmp_path):
    core = (
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">'
        "<dc:title>Quarterly report</dc:title><dc:creator>A. Author</dc:creator>"
        "<dcterms:created>2024-03-05T09:30:00Z</dcterms:created></cp:coreProperties>"
    )
    path = make_docx(tmp_path, p("x"), core=core)
    props = core_properties(path)
    assert props["dc:title"] == "Quarterly report"
    created = created_date(path)
    assert (created.year, created.month, created.day, created.utcoffset().total_seconds()) == (2024, 3, 5, 0)


def test_no_core_properties(tmp_path):
    path = make_docx(tmp_path, p("x"))
    assert core_properties(path) == {}
    assert created_date(path) is None
