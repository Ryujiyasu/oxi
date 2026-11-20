# SPDX-License-Identifier: MIT OR Apache-2.0
"""Run against the installed package: pip install . && pytest tests/."""

from pathlib import Path
import threading

import pytest

import oxidocs

SAMPLES = Path(__file__).resolve().parents[3] / "integrations" / "gotenberg" / "e2e" / "samples"
SIMPLE = SAMPLES / "simple.docx"          # Apache POI test document
LISTS = SAMPLES / "lists.docx"            # docling test document
UNSUPPORTED = SAMPLES / "unsupported.docx"  # LibreOffice test document with SmartArt Oxi cannot draw
FIXTURES = Path(__file__).resolve().parents[3] / "tests" / "fixtures"
SLIDES = FIXTURES / "multi_slide.pptx"       # three slides
GRADIENT = FIXTURES / "gradalpha_test.pptx"  # a gradient fill, not drawn yet


def test_converts_a_document(tmp_path):
    result = oxidocs.docx_to_pdf(SIMPLE, tmp_path / "out.pdf")
    assert result.pages >= 1
    assert result.output.read_bytes()[:5] == b"%PDF-"
    assert result.bytes == result.output.stat().st_size


def test_strict_reports_content_it_cannot_draw(tmp_path):
    with pytest.raises(oxidocs.UnsupportedContent) as caught:
        oxidocs.docx_to_pdf(UNSUPPORTED, tmp_path / "out.pdf", strict=True)
    assert "smartart" in caught.value.kinds
    assert caught.value.exit_status == 3
    # Without strict the same document still converts.
    assert oxidocs.docx_to_pdf(UNSUPPORTED, tmp_path / "loose.pdf").pages >= 1


def test_missing_and_unreadable_files(tmp_path):
    with pytest.raises(oxidocs.ConversionError):
        oxidocs.docx_to_pdf(tmp_path / "nope.docx", tmp_path / "out.pdf")
    junk = tmp_path / "junk.docx"
    junk.write_bytes(b"not a zip")
    with pytest.raises(oxidocs.ConversionError) as caught:
        oxidocs.docx_to_pdf(junk, tmp_path / "out.pdf")
    assert not isinstance(caught.value, oxidocs.UnsupportedContent)


def test_converter_serves_many_documents_concurrently(tmp_path):
    results = {}
    with oxidocs.Converter(threads=2) as converter:
        def run(i, src):
            results[i] = converter.convert(src, tmp_path / f"{i}.pdf")
        threads = [threading.Thread(target=run, args=(i, src)) for i, src in enumerate([SIMPLE, LISTS] * 3)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
    assert sorted(results) == list(range(6))
    assert all(r.output.read_bytes()[:5] == b"%PDF-" for r in results.values())


def test_converts_a_presentation_one_page_per_slide(tmp_path):
    result = oxidocs.pptx_to_pdf(SLIDES, tmp_path / "deck.pdf")
    assert result.pages == 3
    assert result.output.read_bytes()[:5] == b"%PDF-"


def test_strict_reports_slide_content_it_cannot_draw(tmp_path):
    with pytest.raises(oxidocs.UnsupportedContent) as caught:
        oxidocs.pptx_to_pdf(GRADIENT, tmp_path / "out.pdf", strict=True)
    assert "gradient-fill" in caught.value.kinds
    assert not (tmp_path / "out.pdf").exists()


def test_converter_takes_documents_and_presentations(tmp_path):
    with oxidocs.Converter(threads=2) as converter:
        doc = converter.convert(SIMPLE, tmp_path / "doc.pdf")
        deck = converter.convert(SLIDES, tmp_path / "deck.pdf")
    assert doc.pages >= 1 and deck.pages == 3


def _docx_with_page_breaks(path, texts):
    """A minimal .docx: one paragraph per text, a page break between them."""
    import zipfile

    body = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'.join(
        f"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>" for t in texts
    )
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml",
            '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            "</Types>")
        z.writestr("_rels/.rels",
            '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            "</Relationships>")
        z.writestr("word/document.xml",
            f'<?xml version="1.0"?><w:document xmlns:w="{w}"><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>')
    return path


def test_docx_to_docling_gives_each_item_its_page(tmp_path):
    src = _docx_with_page_breaks(tmp_path / "three.docx", ["First page", "Second page", "Third page"])
    doc = oxidocs.docx_to_docling(src)
    assert doc["schema_name"] == "DoclingDocument"
    assert len(doc["pages"]) == 3
    # A4 in points.
    assert round(doc["pages"]["1"]["size"]["width"]) == 595
    pages = {t["text"]: [p["page_no"] for p in t["prov"]] for t in doc["texts"] if t["text"]}
    assert pages == {"First page": [1], "Second page": [2], "Third page": [3]}
    box = next(t for t in doc["texts"] if t["text"] == "Second page")["prov"][0]["bbox"]
    assert box["coord_origin"] == "TOPLEFT" and 0 < box["l"] < box["r"] and 0 < box["t"] < box["b"]


def test_docx_to_docling_reports_unreadable_files(tmp_path):
    junk = tmp_path / "junk.docx"
    junk.write_bytes(b"not a zip")
    with pytest.raises(oxidocs.ConversionError):
        oxidocs.docx_to_docling(junk)


def test_ships_the_notices_for_the_ported_equation_converter():
    from importlib.resources import files

    notices = files("oxidocs").joinpath("THIRD_PARTY_NOTICES.md").read_text(encoding="utf-8")
    assert "Copyright The Docling Contributors" in notices and "dwml" in notices
