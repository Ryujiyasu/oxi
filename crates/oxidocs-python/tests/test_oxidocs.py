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
