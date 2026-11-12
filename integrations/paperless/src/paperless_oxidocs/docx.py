# SPDX-License-Identifier: MIT OR Apache-2.0
"""Plain text and core properties read straight from a .docx package.

Paperless indexes the text a parser returns. Reading it from the document's
own XML (rather than back out of the PDF) keeps words whole across line
breaks and keeps reading order, as Tika's text did.
"""

from __future__ import annotations

import datetime
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
_STORY = re.compile(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")
_ORDER = {"document": 0, "header": 1, "footer": 2, "footnotes": 3, "endnotes": 4}


def _paragraph_text(p: ET.Element) -> str:
    out: list[str] = []

    def walk(el: ET.Element) -> None:
        for child in el:
            tag = child.tag
            if tag == MC_FALLBACK:
                continue  # a second copy of what mc:Choice holds
            if tag == W + "t":
                out.append(child.text or "")
            elif tag == W + "tab":
                out.append("\t")
            elif tag in (W + "br", W + "cr"):
                out.append("\n")
            elif tag == W + "p":
                continue  # nested paragraphs (text boxes) are read on their own
            elif tag == W + "delText":
                continue  # tracked deletion
            else:
                walk(child)

    walk(p)
    return "".join(out)


def _story_text(xml: bytes) -> list[str]:
    root = ET.fromstring(xml)
    # Drop every mc:Fallback subtree so text boxes are read once.
    for parent in list(root.iter()):
        for child in list(parent):
            if child.tag == MC_FALLBACK:
                parent.remove(child)
    lines = []
    for p in root.iter(W + "p"):
        text = _paragraph_text(p)
        if text.strip():
            lines.append(text.rstrip())
    return lines


def docx_text(path: Path) -> str:
    """The document's text: body (tables and text boxes included), then
    headers, footers, footnotes and endnotes, one paragraph per line."""
    with zipfile.ZipFile(path) as z:
        parts = [n for n in z.namelist() if _STORY.search(n)]
        parts.sort(key=lambda n: (_ORDER.get(re.sub(r"\d+$", "", n[5:-4]), 9), n))
        lines: list[str] = []
        # A header or footer that several sections repeat reads once.
        seen_hf: set[str] = set()
        for name in parts:
            try:
                story = _story_text(z.read(name))
            except ET.ParseError:
                continue
            if "/header" in name or "/footer" in name:
                story = [line for line in story if not (line in seen_hf or seen_hf.add(line))]
            lines.extend(story)
    return "\n".join(lines).strip()


_CORE = {
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
}


def core_properties(path: Path) -> dict[str, str]:
    """docProps/core.xml as {"dc:title": ..., "dcterms:created": ...}."""
    try:
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("docProps/core.xml"))
    except (KeyError, zipfile.BadZipFile, ET.ParseError, OSError):
        return {}
    by_uri = {uri: prefix for prefix, uri in _CORE.items()}
    props = {}
    for el in root:
        uri, _, local = el.tag[1:].partition("}")
        prefix = by_uri.get(uri)
        if prefix and el.text and el.text.strip():
            props[f"{prefix}:{local}"] = el.text.strip()
    return props


def created_date(path: Path) -> datetime.datetime | None:
    """dcterms:created as an aware datetime, or None."""
    value = core_properties(path).get("dcterms:created")
    if not value:
        return None
    try:
        parsed = datetime.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=datetime.timezone.utc)
    # Word writes 1601-01-01 or similar for "unset".
    return parsed if parsed.year >= 1980 else None
