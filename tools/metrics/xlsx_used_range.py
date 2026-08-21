# -*- coding: utf-8 -*-
"""Does Oxi draw the same range of the sheet that Excel hands over?

Every comparison in the gate starts at the top-left of a range: Excel's
picture is of `UsedRange`, and Oxi draws what it works out for itself. A row
or column of difference at the near end shifts the whole sheet against itself
and reads as a total mismatch — data_A28 sat at 0.57 for exactly that reason.

This asks Excel for the range and asks the renderer for its own, and reports
where they differ.

    python tools\\metrics\\xlsx_used_range.py <dir-or-file.xlsx> [--limit N]
"""
import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
OUT = REPO / "pipeline_data" / "xlsx_used_range.json"


def excel_ranges(paths):
    """First and last row and column of each sheet's used range."""
    import win32com.client

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    held = {}
    try:
        for path in paths:
            try:
                book = excel.Workbooks.Open(str(path.resolve()), 0, True)
            except Exception as bad:
                print(f"  {path.stem[:40]:42s} Excel would not open it: {str(bad)[:60]}")
                continue
            try:
                sheet = book.Worksheets(1)
                used = sheet.UsedRange
                held[path.stem] = (int(used.Row), int(used.Column),
                                   int(used.Row) + int(used.Rows.Count) - 1,
                                   int(used.Column) + int(used.Columns.Count) - 1)
            finally:
                book.Close(False)
    finally:
        excel.Quit()
    return held


def our_range(path):
    environment = dict(os.environ, OXI_XLSX_DUMP_ROWS="1", OXI_XLSX_DUMP_COLUMNS="1")
    done = subprocess.run(
        [str(RENDERER), str(path), str(Path(os.environ.get("TEMP", ".")) / "_range.png"), "96"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=600, env=environment)
    rows, columns = [], []
    for line in done.stdout.splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "row":
            rows.append(int(parts[1]))
        if len(parts) == 4 and parts[0] == "column":
            columns.append(int(parts[1]))
    if not rows or not columns:
        return None
    # The dump counts rows the way the sheet does and columns from zero.
    return (min(rows), min(columns) + 1, max(rows), max(columns) + 1)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("target", type=Path)
    parser.add_argument("--limit", type=int)
    args = parser.parse_args()

    paths = sorted(args.target.glob("*.xlsx")) if args.target.is_dir() else [args.target]
    paths = [path for path in paths if not path.name.startswith("~$")]
    if args.limit:
        paths = paths[: args.limit]

    theirs = excel_ranges(paths)
    agree, held = 0, {}
    for path in paths:
        want = theirs.get(path.stem)
        if want is None:
            continue
        got = our_range(path)
        held[path.stem] = {"excel": want, "oxi": got}
        if got == want:
            agree += 1
            continue
        where = []
        for index, name in enumerate(["first row", "first column", "last row", "last column"]):
            if got is None or got[index] != want[index]:
                where.append(f"{name} {want[index]} vs {got[index] if got else '-'}")
        print(f"  {path.stem[:44]:46s} {', '.join(where)}")

    print(f"\n{agree} of {len(theirs)} workbooks draw the range Excel hands over")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(held, indent=1), encoding="utf-8")
    print(f"written to {OUT}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
