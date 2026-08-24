# -*- coding: utf-8 -*-
"""Which characters does the corpus actually stack?

A stacked cell is drawn through two faces — the upright one for what Excel
leaves standing, the turned one for the marks it lays on their side — so the
class in `turned_in_a_stack` has to cover whatever the corpus actually puts in
a stacked cell. This counts them.

The sheets are read with a pull parser, not a regular expression: a 20MB
worksheet answers a `<c ...>(.*?)</c>` pattern slowly enough to look hung.

    python tools\\metrics\\_xlsx_stack_census.py
"""

from __future__ import annotations

import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
BOOKS = REPO / "tools" / "golden-test" / "documents" / "xlsx"
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def stacked_styles(styles: bytes) -> set[int]:
    """The cellXfs indices whose alignment stands the characters up."""
    out: set[int] = set()
    root = ET.fromstring(styles)
    block = root.find(f"{MAIN}cellXfs")
    if block is None:
        return out
    for at, xf in enumerate(block):
        alignment = xf.find(f"{MAIN}alignment")
        if alignment is not None and alignment.get("textRotation") == "255":
            out.add(at)
    return out


def shared(strings: bytes) -> list[str]:
    root = ET.fromstring(strings)
    return ["".join(t.text or "" for t in item.iter(f"{MAIN}t")) for item in root]


def main() -> int:
    tally: Counter[str] = Counter()
    books: dict[str, set[str]] = {}
    for book in sorted(BOOKS.glob("*.xlsx")):
        try:
            zipped = zipfile.ZipFile(book)
        except zipfile.BadZipFile:
            continue
        names = set(zipped.namelist())
        if "xl/styles.xml" not in names:
            continue
        stacked = stacked_styles(zipped.read("xl/styles.xml"))
        if not stacked:
            continue
        table = (shared(zipped.read("xl/sharedStrings.xml"))
                 if "xl/sharedStrings.xml" in names else [])
        for sheet in sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)):
            with zipped.open(sheet) as body:
                for event, node in ET.iterparse(body, events=("end",)):
                    if node.tag != f"{MAIN}c":
                        continue
                    style = node.get("s")
                    if style is None or int(style) not in stacked:
                        node.clear()
                        continue
                    said = node.find(f"{MAIN}v")
                    if node.get("t") == "s" and said is not None:
                        try:
                            words = table[int(said.text or "-1")]
                        except (ValueError, IndexError):
                            words = ""
                    else:
                        words = "".join(t.text or "" for t in node.iter(f"{MAIN}t"))
                        if not words and said is not None:
                            words = said.text or ""
                    for letter in words:
                        tally[letter] += 1
                        books.setdefault(letter, set()).add(book.stem)
                    node.clear()
    print(f"  {len(tally)} distinct characters stacked, {sum(tally.values())} in all")
    print("  count  books  char")
    for letter, count in tally.most_common():
        shown = letter.encode("unicode_escape").decode("ascii")
        print(f"  {count:>5}  {len(books[letter]):>5}  {shown}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
