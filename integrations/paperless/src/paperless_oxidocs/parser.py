# SPDX-License-Identifier: MIT OR Apache-2.0
"""Paperless-ngx parser for .docx, converted locally by the Oxi engine.

Paperless-ngx sends Office documents to two external services: Tika for
their text and Gotenberg (LibreOffice) for the PDF it displays. This parser
claims .docx instead and does both inside the Paperless container:

- the PDF comes from the bundled oxidocs engine (``pip install oxidocs``),
  turned into PDF/A by the container's own ocrmypdf when the configured
  output type asks for it;
- the text, date and metadata come from the document's own XML.

A document holding content Oxi does not draw (oxidocs ``strict`` mode) goes
to the built-in Tika parser when Tika is enabled, so nothing converts worse
than before; without Tika it is converted as well as Oxi can.
"""

from __future__ import annotations

import logging
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import TYPE_CHECKING, Self

import oxidocs
from django.conf import settings

from . import __version__
from .docx import core_properties, created_date, docx_text

if TYPE_CHECKING:
    import datetime
    from types import TracebackType

    from paperless.parsers import MetadataEntry, ParserContext

logger = logging.getLogger("paperless.parsing.oxidocs")

DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


class OxiDocxParser:
    name = "Oxi DOCX Parser"
    version = __version__
    author = "Oxi contributors"
    url = "https://github.com/Ryujiyasu/oxi/tree/main/integrations/paperless"

    @classmethod
    def supported_mime_types(cls) -> dict[str, str]:
        return {DOCX: ".docx"}

    @classmethod
    def score(cls, mime_type: str, filename: str, path: Path | None = None) -> int | None:
        # Above the built-in Tika parser's 10, so .docx comes here whether or
        # not Tika is enabled; other Office formats stay with Tika.
        return 15 if mime_type == DOCX else None

    @property
    def can_produce_archive(self) -> bool:
        return False

    @property
    def requires_pdf_rendition(self) -> bool:
        # A browser cannot show .docx: the PDF is always produced.
        return True

    # -- lifecycle ---------------------------------------------------------

    def __init__(self, logging_group: object = None) -> None:
        settings.SCRATCH_DIR.mkdir(parents=True, exist_ok=True)
        self._tempdir = Path(tempfile.mkdtemp(prefix="paperless-oxidocs-", dir=settings.SCRATCH_DIR))
        self._logging_group = logging_group
        self._context: ParserContext | None = None
        self._text: str = ""
        self._date: datetime.datetime | None = None
        self._archive_path: Path | None = None

    def __enter__(self) -> Self:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None,
    ) -> None:
        shutil.rmtree(self._tempdir, ignore_errors=True)

    def configure(self, context: ParserContext) -> None:
        self._context = context

    # -- parsing -----------------------------------------------------------

    def parse(self, document_path: Path, mime_type: str, *, produce_archive: bool = True) -> None:
        from documents.parsers import ParseError

        pdf = self._tempdir / "convert.pdf"
        try:
            result = oxidocs.docx_to_pdf(document_path, pdf, strict=True, timeout=settings.CELERY_TASK_TIME_LIMIT)
            logger.info("Oxi converted %s: %d pages in %.0f ms", document_path.name, result.pages, result.ms)
        except oxidocs.UnsupportedContent as unsupported:
            if self._parse_with_tika(document_path, mime_type, unsupported.kinds):
                return
            logger.warning(
                "%s holds content Oxi does not draw (%s) and Tika is off: converting it as well as Oxi can",
                document_path.name,
                ", ".join(unsupported.kinds),
            )
            try:
                oxidocs.docx_to_pdf(document_path, pdf, timeout=settings.CELERY_TASK_TIME_LIMIT)
            except oxidocs.ConversionError as err:
                raise ParseError(f"Oxi could not convert {document_path.name}: {err}") from err
        except oxidocs.ConversionError as err:
            raise ParseError(f"Oxi could not convert {document_path.name}: {err}") from err

        try:
            self._text = docx_text(document_path)
        except Exception as err:  # the PDF still stands without searchable text
            logger.warning("Could not read the text of %s: %s", document_path.name, err)
        self._date = created_date(document_path)
        self._archive_path = self._to_output_type(pdf)

    def _parse_with_tika(self, document_path: Path, mime_type: str, kinds: list[str]) -> bool:
        """Hand a document Oxi does not fully draw to the built-in Tika
        parser; False when Tika is not enabled."""
        if not getattr(settings, "TIKA_ENABLED", False):
            return False
        from paperless.parsers.tika import TikaDocumentParser

        logger.info("%s holds content Oxi does not draw (%s): handing it to Tika", document_path.name, ", ".join(kinds))
        with TikaDocumentParser(self._logging_group) as tika:
            if self._context is not None:
                tika.configure(self._context)
            tika.parse(document_path, mime_type)
            self._text = tika.get_text()
            self._date = tika.get_date()
            archive = tika.get_archive_path()
            if archive is not None:
                # Tika's working directory goes when its context closes.
                self._archive_path = self._tempdir / "tika.pdf"
                shutil.copyfile(archive, self._archive_path)
        return True

    def _to_output_type(self, pdf: Path) -> Path:
        """The PDF in the configured output type (PAPERLESS_OCR_OUTPUT_TYPE or
        the application setting): PDF/A through ocrmypdf, as Paperless makes
        its other archives. Oxi's PDF already carries its text, so no OCR."""
        from paperless.config import OutputTypeConfig

        output_type = str(OutputTypeConfig().output_type)
        if output_type == "pdf":
            return pdf
        out = self._tempdir / "archive.pdf"
        command = ["ocrmypdf", "--skip-text", "--output-type", output_type, "--quiet", str(pdf), str(out)]
        try:
            subprocess.run(command, check=True, capture_output=True, timeout=settings.CELERY_TASK_TIME_LIMIT)
            return out
        except (OSError, subprocess.SubprocessError) as err:
            detail = getattr(err, "stderr", b"") or b""
            logger.warning(
                "ocrmypdf could not make %s of %s (%s); keeping the plain PDF",
                output_type,
                pdf.name,
                detail.decode("utf-8", "replace").strip()[-300:] or err,
            )
            return pdf

    # -- results -----------------------------------------------------------

    def get_text(self) -> str:
        return self._text or ""

    def get_date(self) -> datetime.datetime | None:
        return self._date

    def get_archive_path(self) -> Path | None:
        return self._archive_path

    def get_thumbnail(self, document_path: Path, mime_type: str) -> Path:
        from documents.parsers import make_thumbnail_from_pdf

        if self._archive_path is None:
            self.parse(document_path, mime_type)
        assert self._archive_path is not None
        return make_thumbnail_from_pdf(self._archive_path, self._tempdir, self._logging_group)

    def get_page_count(self, document_path: Path, mime_type: str) -> int | None:
        if self._archive_path is None:
            return None
        from paperless.parsers.utils import get_page_count_for_pdf

        return get_page_count_for_pdf(self._archive_path, log=logger)

    def extract_metadata(self, document_path: Path, mime_type: str) -> list[MetadataEntry]:
        uris = {
            "dc": "http://purl.org/dc/elements/1.1/",
            "dcterms": "http://purl.org/dc/terms/",
            "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        }
        try:
            props = core_properties(document_path)
        except Exception:  # must never raise
            return []
        entries: list[MetadataEntry] = []
        for qualified, value in sorted(props.items()):
            prefix, _, key = qualified.partition(":")
            entries.append({"namespace": uris.get(prefix, ""), "prefix": prefix, "key": key, "value": value})
        return entries
