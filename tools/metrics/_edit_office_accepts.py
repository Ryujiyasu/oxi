# -*- coding: utf-8 -*-
r"""Does Office open what our editor wrote?

`oxi-roundtrip` proves the writer kept everything the IR models. It cannot
prove the file still opens: a workbook whose parts we patched into an
inconsistent state parses back perfectly through our own reader and is refused
— or silently repaired — by Excel. That is the worst way for an editor to
fail, because nothing before this notices it.

So this points Office at the saved copies. A file Excel has to repair is NOT
accepted: `CorruptLoad=xlNormalLoad` makes it raise rather than quietly mend
the file, so a repair shows up as a refusal instead of passing as a success.

    oxi-roundtrip <corpus> --keep C:\tmp\edited --quiet
    python tools\metrics\_edit_office_accepts.py C:\tmp\edited
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import win32com.client

# xlNormalLoad: open as-is. The other two (xlRepairFile, xlExtractData) are
# what Excel falls back to on its own when a file is damaged, and either would
# turn a broken write into a passing one.
NORMAL_LOAD = 0


def workbooks(excel, files: list[Path]) -> list[tuple[Path, str]]:
    trouble = []
    for path in files:
        book = None
        try:
            book = excel.Workbooks.Open(
                str(path), UpdateLinks=0, ReadOnly=True, CorruptLoad=NORMAL_LOAD
            )
            # Touching a sheet forces Excel to actually read the parts rather
            # than hand back a lazily opened shell.
            _ = book.Worksheets(1).Name
        except Exception as refused:
            trouble.append((path, str(refused).split("',")[0][:90]))
        finally:
            if book is not None:
                try:
                    book.Close(SaveChanges=False)
                except Exception:
                    pass
    return trouble


def documents(word, files: list[Path]) -> list[tuple[Path, str]]:
    trouble = []
    for path in files:
        doc = None
        try:
            doc = word.Documents.Open(
                str(path), ConfirmConversions=False, ReadOnly=True,
                AddToRecentFiles=False, Visible=False,
            )
            _ = doc.Paragraphs.Count
        except Exception as refused:
            trouble.append((path, str(refused).split("',")[0][:90]))
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass
    return trouble


def decks(power, files: list[Path]) -> list[tuple[Path, str]]:
    trouble = []
    for path in files:
        deck = None
        try:
            deck = power.Presentations.Open(
                str(path), ReadOnly=True, Untitled=False, WithWindow=False
            )
            _ = deck.Slides.Count
        except Exception as refused:
            trouble.append((path, str(refused).split("',")[0][:90]))
        finally:
            if deck is not None:
                try:
                    deck.Close()
                except Exception:
                    pass
    return trouble


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("where", help="directory of files the editor wrote")
    parser.add_argument("--limit", type=int, default=0)
    # A file Office refuses is only OUR doing if Office took the original.
    # Two of the deck fixtures are hand-made and PowerPoint has never opened
    # them; counting those as damage would put a permanent red mark on a
    # metric that is meant to mean something.
    parser.add_argument("--originals", default="", help="where the inputs came from")
    args = parser.parse_args()
    where = Path(args.where)
    sheets = sorted(p for p in where.glob("*.xls*") if not p.name.startswith("~$"))
    docs = sorted(p for p in where.glob("*.docx") if not p.name.startswith("~$"))
    slides = sorted(p for p in where.glob("*.pptx") if not p.name.startswith("~$"))
    if args.limit:
        sheets = sheets[: args.limit]
        docs = docs[: args.limit]
        slides = slides[: args.limit]

    refused: list[tuple[Path, str]] = []
    if sheets:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        try:
            refused += workbooks(excel, sheets)
        finally:
            try:
                excel.Quit()
            except Exception:
                pass
    if docs:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            refused += documents(word, docs)
        finally:
            # Word's RPC endpoint sometimes goes before we ask it to. Losing
            # the whole run's findings to that would be a poor trade.
            try:
                word.Quit()
            except Exception:
                pass

    if slides:
        # PowerPoint has no CorruptLoad of its own: a deck it quietly mends
        # still counts as opened here, which is worth knowing when reading
        # this number.
        power = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            refused += decks(power, slides)
        finally:
            try:
                power.Quit()
            except Exception:
                pass

    total = len(sheets) + len(docs) + len(slides)
    ours, theirs = refused, []
    if refused and args.originals:
        source = Path(args.originals)
        ours, theirs = [], []
        again = [(source / path.name) for path, _ in refused if (source / path.name).exists()]
        was = {p.name for p, _ in check(again)}
        for path, why in refused:
            (theirs if path.name in was else ours).append((path, why))
    print(f"  {total} file(s) the editor wrote: {total - len(ours)} opened as they are")
    for path, why in ours:
        print(f"    XX  {path.name}  {why}")
    for path, _ in theirs:
        print(f"    --  {path.name}  (Office will not open the ORIGINAL either)")
    return 0


def check(files: list[Path]) -> list[tuple[Path, str]]:
    """Which of these Office refuses, whatever they are."""
    out: list[tuple[Path, str]] = []
    for kind, opener, app in (
        ("*.xls*", workbooks, "Excel.Application"),
        ("*.docx", documents, "Word.Application"),
        ("*.pptx", decks, "PowerPoint.Application"),
    ):
        held = [p for p in files if p.match(kind)]
        if not held:
            continue
        office = win32com.client.DispatchEx(app)
        try:
            if app != "PowerPoint.Application":
                office.Visible = False
            out += opener(office, held)
        finally:
            office.Quit()
    return out


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
