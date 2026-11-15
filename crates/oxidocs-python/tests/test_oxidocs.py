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
