# -*- coding: utf-8 -*-
r"""What does Excel put in the cells you drag the fill handle over?

The handle is several rules wearing one gesture, and which one Excel picks
depends on what was selected. Two of them are easy to get wrong by reasoning
rather than looking:

  * a single number — is it repeated, or counted up?
  * numbers that are not evenly spaced — does the series carry on from the
    LAST gap, or along a line fitted through all of them? For 1, 2, 4 those
    give 5.5 and 5.333, and there is no way to tell which from first
    principles.

So this asks. Each arm writes a few cells, drags the handle down over the ones
below, and reads back what landed there.

    python tools\metrics\_xlsx_fill_series.py
"""

from __future__ import annotations

import sys

import win32com.client

# Each arm: what to seed the column with, and how many cells to pull it over.
ARMS: list[tuple[str, list, int]] = [
    ("one number", [5], 3),
    ("two numbers", [1, 2], 4),
    ("a bigger step", [10, 20, 30], 2),
    ("a step going down", [10, 8], 3),
    ("unevenly spaced", [1, 2, 4], 2),
    ("more unevenly spaced", [1, 4, 5], 2),
    ("text", ["total"], 2),
    ("text ending in digits", ["Item 1"], 3),
    ("text with leading zeros", ["A001"], 2),
    ("two unrelated words", ["a", "b"], 4),
    ("a weekday", ["Mon"], 3),
    ("a month", ["January"], 2),
    ("a Japanese weekday", ["月"], 3),
    ("a date", ["2026/01/31"], 2),
    ("a number then text", [1, "a"], 3),
    # A number sitting inside a mixed block counts up, though the same number
    # alone would repeat. These pin what the step is and where it comes from.
    ("text then a number", ["a", 1], 4),
    ("a number that is not one", [2, "a"], 4),
    ("a big number in a block", [100, "a"], 4),
    ("two numbers in a block", [2, 4, "a"], 4),
    ("numbers either side of text", [1, "a", 9], 4),
    ("two strides of text", [1, "a", "b"], 4),
    ("a weekday inside a block", ["Mon", 1], 4),
    # Does the trailing-digit rule apply to a run inside a block, or only to a
    # cell selected on its own?
    ("numbered text in a block", ["Item 1", "a"], 4),
    ("numbered text after text", ["a", "Item 1"], 4),
    ("two numbered texts", ["Item 1", "Item 2"], 4),
    # Which lists does Excel know? Each of these is a single cell pulled down
    # three: one that continues is a list, one that repeats is not.
    ("short weekday", ["Sun"], 3),
    ("long weekday", ["Sunday"], 3),
    ("short month", ["Jan"], 3),
    ("Japanese long weekday", ["日曜日"], 3),
    ("Japanese old month", ["睦月"], 3),
    ("Japanese numbered month", ["1月"], 3),
    ("quarter", ["第1四半期"], 3),
    ("stems", ["甲"], 3),
    ("iroha", ["い"], 3),
    ("zodiac", ["子"], 3),
    ("a bare number as text", ["1"], 3),
    # Two members of one list: does the run carry the stride between them?
    ("two weekdays in a row", ["Sun", "Mon"], 4),
    ("two weekdays two apart", ["Mon", "Wed"], 4),
    ("two weekdays backwards", ["Wed", "Mon"], 4),
    ("a weekday and a month", ["Mon", "Jan"], 4),
    ("numbered text sharing no prefix", ["Item 1", "Row 5"], 4),
    # Dates are numbers wearing a format. A lone number repeats; does a lone
    # date? And do two dates carry the gap between them, or always a day?
    ("one date", ["2026/01/30"], 3),
    ("one date at a month end", ["2026/01/31"], 3),
    ("two dates a day apart", ["2026/01/30", "2026/01/31"], 3),
    ("two dates a week apart", ["2026/01/05", "2026/01/12"], 3),
    ("two dates a month apart", ["2026/01/31", "2026/02/28"], 3),
    ("a date in a block", ["2026/01/30", "a"], 4),
    ("a time", ["10:30"], 3),
]


def main() -> int:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        at = 1
        for what, seed, pull in ARMS:
            top = at
            for offset, value in enumerate(seed):
                sheet.Cells(top + offset, 1).Value = value
            last = top + len(seed) - 1
            source = sheet.Range(sheet.Cells(top, 1), sheet.Cells(last, 1))
            whole = sheet.Range(sheet.Cells(top, 1), sheet.Cells(last + pull, 1))
            try:
                source.AutoFill(Destination=whole)
            except Exception as error:  # noqa: BLE001 — report, do not stop
                print(f"  {what:<24} would not fill: {error}")
                at = last + pull + 3
                continue
            got = [sheet.Cells(top + n, 1).Text
                   for n in range(len(seed) + pull)]
            print(f"  {what:<24}{str(seed):<22} -> {got}")
            # A date is a number wearing a format, and the display truncates in
            # a narrow column. The serial underneath is what the rule is about.
            if any(one.startswith("#") for one in got):
                raw = [sheet.Cells(top + n, 1).Value2
                       for n in range(len(seed) + pull)]
                print(f"  {'':<24}{'':<22}    serials {raw}")
            at = last + pull + 3
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
