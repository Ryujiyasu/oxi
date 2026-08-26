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


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("where", help="directory of files the editor wrote")
    parser.add_argument("--limit", type=int, default=0)
    args = parser.parse_args()
    where = Path(args.where)
    sheets = sorted(p for p in where.glob("*.xls*") if not p.name.startswith("~$"))
    docs = sorted(p for p in where.glob("*.docx") if not p.name.startswith("~$"))
    if args.limit:
        sheets, docs = sheets[: args.limit], docs[: args.limit]

    refused: list[tuple[Path, str]] = []
    if sheets:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        try:
            refused += workbooks(excel, sheets)
        finally:
            excel.Quit()
    if docs:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            refused += documents(word, docs)
        finally:
            word.Quit()

    total = len(sheets) + len(docs)
    print(f"  {total} file(s) the editor wrote: {total - len(refused)} opened as they are")
    for path, why in refused:
        print(f"    XX  {path.name}  {why}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
