# -*- coding: utf-8 -*-
r"""Does an edit arrive where it was aimed, and nowhere else?

The three layers so far all rest on an edit that asks for nothing to change.
That proves the writer breaks nothing; it does not prove an edit ARRIVES. A
writer that dropped every edit on the floor would pass all three.

So this makes a real change. `oxi-roundtrip --sentinel` writes a mark into one
cell and says which cell it aimed at; this opens the original and the edited
copy side by side in Excel and compares every cell of the used range. Exactly
one must differ, it must be the cell that was aimed at, and it must hold the
mark.

    oxi-roundtrip <corpus> --limit 12 --quiet --sentinel OXIMARK ^
        --keep C:\tmp\aimed_xlsx > C:\tmp\aimed.txt
    python tools\metrics\_edit_lands_where_aimed.py C:\tmp\aimed_xlsx C:\tmp\aimed.txt

Excel is the judge rather than our own reader, because the question is whether
the change reached the file as Excel understands it.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
ORIGINALS = REPO / "tools" / "golden-test" / "documents" / "xlsx"
AIMED = re.compile(r"^\s*aimed (\S+) sheet (\d+) row (\d+) col (\d+)\s*$")


def seen(sheet) -> dict[tuple[int, int], str]:
    """Every cell Excel shows on this sheet, by row and column."""
    used = sheet.UsedRange
    values = used.Value
    if values is None:
        return {}
    top, left = used.Row, used.Column
    if not isinstance(values, tuple):
        return {(top, left): str(values)}
    held = {}
    for down, line in enumerate(values):
        if not isinstance(line, tuple):
            line = (line,)
        for across, one in enumerate(line):
            if one is not None:
                held[(top + down, left + across)] = str(one)
    return held


def read_all(excel, path: Path) -> dict[int, dict[tuple[int, int], str]] | None:
    """Every sheet's cells as Excel shows them, keyed by sheet number."""
    book = None
    try:
        book = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
        if book is None:
            return None
        return {at: seen(book.Worksheets(at)) for at in range(1, book.Worksheets.Count + 1)}
    except Exception:
        return None
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("where", help="directory of files the editor wrote")
    parser.add_argument("aimed", help="the tool's output, saying where it aimed")
    parser.add_argument("--mark", default="OXIMARK")
    args = parser.parse_args()

    targets: dict[str, tuple[int, int, int]] = {}
    for line in Path(args.aimed).read_text(encoding="utf-8", errors="ignore").splitlines():
        found = AIMED.match(line)
        if found:
            targets[found.group(1)] = tuple(int(g) for g in found.groups()[1:])

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    landed, wrong = 0, []
    try:
        for name, (sheet_at, row, col) in sorted(targets.items()):
            edited = Path(args.where) / name
            before = ORIGINALS / name
            if not edited.exists() or not before.exists():
                continue
            # One at a time: Excel will not hold two workbooks of the same
            # name at once, and the edited copy keeps the name it came with.
            was = read_all(excel, before)
            now = read_all(excel, edited)
            if was is None or now is None:
                wrong.append((name, "Excel would not open one of the two"))
                continue
            # The editor counts sheets from zero and columns from zero, but
            # ROWS from one — `Row.index` carries the `<row r="...">` the file
            # states. Excel counts all three from one. Reading that wart the
            # obvious way made all twelve arms look wrong when every one of
            # them had landed exactly right.
            want = (row, col + 1)
            moved = {
                (at, cell)
                for at in set(was) | set(now)
                for cell in set(was.get(at, {})) | set(now.get(at, {}))
                if was.get(at, {}).get(cell) != now.get(at, {}).get(cell)
            }
            aimed_at = (sheet_at + 1, want)
            if moved == {aimed_at} and now.get(sheet_at + 1, {}).get(want) == args.mark:
                landed += 1
            elif not moved:
                wrong.append((name, "the mark never arrived"))
            elif moved != {aimed_at}:
                others = sorted(moved - {aimed_at})[:3]
                wrong.append((name, f"{len(moved)} cell(s) moved, including {others}"))
            else:
                held = now.get(sheet_at + 1, {}).get(want)
                wrong.append((name, f"holds {held!r}, not the mark"))
    finally:
        excel.Quit()

    print(f"  {len(targets)} aimed edit(s): {landed} arrived where they were aimed, alone")
    for name, why in wrong:
        print(f"    !!  {name}  {why}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
