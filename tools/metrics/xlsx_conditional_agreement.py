# -*- coding: utf-8 -*-
"""Whether a conditionally formatted sheet looks the way Excel makes it look.

Every other xlsx metric here asks about geometry. This one asks about the one
part of a sheet whose look is not written down anywhere in the file: a
conditional rule lives on the sheet, not on the cell, so the only way to know
what a cell wears is to run the rule against the value.

Excel will answer the same question directly. `Range.DisplayFormat` returns the
format a person actually sees — the cell's own format with every conditional
rule already applied — so the comparison is exact rather than a pixel guess. A
cell is conditionally formatted, to Excel, when its `DisplayFormat` differs
from its own format.

    python tools/metrics/xlsx_conditional_agreement.py tools/golden-test/documents/xlsx

Two things can disagree, and they are counted apart: WHICH cells a rule catches
(a wrong condition), and WHAT it puts on them (a wrong dxf).
"""
import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
DUMPER = REPO / "target" / "release" / "examples" / "_conditional_dump.exe"

XL_NONE = -4142
# One workbook's worth of questions. `DisplayFormat` is a per-cell round trip
# through COM, and a rule stated over a whole column would otherwise hold the
# run for an hour on a single sheet.
CAP = 4000


def bgr(colour) -> str:
    """Excel hands colours back as BGR integers; the file writes RGB hex."""
    value = int(colour)
    return "%02X%02X%02X" % (value & 0xFF, (value >> 8) & 0xFF, (value >> 16) & 0xFF)


def oxi_says(path: Path):
    """Ranges and hits, per sheet, as this engine reads them."""
    # The dumper writes sheet names, which on this corpus are Japanese. Without
    # being told, Python decodes a child's output with the console codepage
    # (cp932 here) and throws on the first UTF-8 byte — which loses the whole
    # workbook silently, and lost 5 of 16 on the first run of this script.
    run = subprocess.run([str(DUMPER), str(path)], capture_output=True, text=True,
                         encoding="utf-8", errors="replace")
    ranges, hits = {}, {}
    for line in (run.stdout or "").splitlines():
        try:
            row = json.loads(line)
        except ValueError:
            continue
        if row["kind"] == "range":
            ranges.setdefault(row["sheet"], []).append(
                (row["top"], row["left"], row["bottom"], row["right"]))
        elif row["kind"] == "hit":
            # A later hit for the same cell is the winning rule.
            hits.setdefault(row["sheet"], {})[(row["row"], row["col"])] = row
    return ranges, hits


def asked_cells(ranges, used_rows, used_cols):
    """The rules' ranges clipped to what the sheet holds, cells counted from zero."""
    want = set()
    for top, left, bottom, right in ranges:
        for row in range(top, min(bottom, used_rows - 1) + 1):
            for col in range(left, min(right, used_cols - 1) + 1):
                want.add((row, col))
                if len(want) > CAP:
                    return want, True
    return want, False


def excel_says(ws, want):
    """What Excel shows on each asked cell, over and above the cell's own format."""
    shown = {}
    for row, col in sorted(want):
        cell = ws.Cells(row + 1, col + 1)
        try:
            worn = cell.DisplayFormat
            own_fill = None if cell.Interior.ColorIndex == XL_NONE else bgr(cell.Interior.Color)
            new_fill = None if worn.Interior.ColorIndex == XL_NONE else bgr(worn.Interior.Color)
            own_font, new_font = bgr(cell.Font.Color), bgr(worn.Font.Color)
            own_bold, new_bold = bool(cell.Font.Bold), bool(worn.Font.Bold)
        except Exception:
            continue
        if (own_fill, own_font, own_bold) == (new_fill, new_font, new_bold):
            continue
        shown[(row, col)] = {
            "bg": new_fill if new_fill != own_fill else None,
            "fg": new_font if new_font != own_font else None,
            "bold": new_bold if new_bold != own_bold else None,
        }
    return shown


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("target", type=Path)
    parser.add_argument("--limit", type=int)
    parser.add_argument("--out", type=Path,
                        default=REPO / "pipeline_data" / "xlsx_conditional_agreement.json")
    args = parser.parse_args()

    if not DUMPER.exists():
        print("build it first: cargo build --release -p oxicells-core "
              "--example _conditional_dump")
        return 2

    sources = sorted(args.target.glob("*.xlsx")) if args.target.is_dir() else [args.target]
    sources = [p for p in sources if not p.name.startswith("~$")]

    # Only the workbooks that carry a rule at all are worth opening in Excel.
    carrying = []
    for source in sources:
        ranges, hits = oxi_says(source)
        if ranges:
            carrying.append((source, ranges, hits))
    if args.limit:
        carrying = carrying[: args.limit]
    print("%d of %d workbooks carry a conditional rule\n" % (len(carrying), len(sources)))

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    report = []
    both = ours_only = theirs_only = 0
    fill_same = fill_off = 0
    try:
        for source, ranges, hits in carrying:
            try:
                wb = excel.Workbooks.Open(str(source.resolve()), 0, True)
            except Exception as error:
                print("  %-40s Excel would not open it: %s"
                      % (source.stem[:40], str(error)[:36]))
                continue
            agreed = missed = extra = 0
            wrong_fill = []
            # Which cells, not just how many: a count says a rule is wrong,
            # the cells say which rule.
            missed_at, extra_at = [], []
            capped = False
            try:
                for at, spans in sorted(ranges.items()):
                    if at >= wb.Worksheets.Count:
                        continue
                    ws = wb.Worksheets(at + 1)
                    used = ws.UsedRange
                    used_rows = used.Row - 1 + used.Rows.Count
                    used_cols = used.Column - 1 + used.Columns.Count
                    want, over = asked_cells(spans, used_rows, used_cols)
                    capped = capped or over
                    shown = excel_says(ws, want)
                    mine = {at_cell: hit for at_cell, hit in hits.get(at, {}).items()
                            if at_cell in want}
                    for cell in set(shown) | set(mine):
                        if cell in shown and cell in mine:
                            agreed += 1
                            # Fill, ink and weight each separately: two rules
                            # over one cell contribute different parts of the
                            # look, so a comparison that only asks about the
                            # fill cannot see one of them go missing.
                            theirs = (shown[cell]["bg"], shown[cell]["fg"], shown[cell]["bold"])
                            got = mine[cell]
                            ours = (got["bg"] or None, got["fg"] or None, got["bold"])
                            # Excel reports only what its DisplayFormat CHANGED
                            # from the cell's own format, so a rule that sets a
                            # colour the cell already wore shows as no change.
                            # A part Excel changed and we do not set is
                            # therefore a real miss; the other direction is not,
                            # and is left alone rather than counted as noise.
                            same = all(
                                a is None or (b is not None and str(a).upper() == str(b).upper())
                                for a, b in zip(theirs, ours))
                            if not same:
                                wrong_fill.append((at, cell[0], cell[1], theirs, ours))
                        elif cell in mine:
                            extra += 1
                            extra_at.append((at, cell[0], cell[1]))
                        else:
                            missed += 1
                            missed_at.append((at, cell[0], cell[1], shown[cell]["bg"]))
            finally:
                wb.Close(False)

            both += agreed
            ours_only += extra
            theirs_only += missed
            fill_off += len(wrong_fill)
            fill_same += agreed - len(wrong_fill)
            report.append({"doc": source.stem, "agreed": agreed, "we_miss": missed,
                           "we_add": extra, "fill_off": len(wrong_fill),
                           "capped": capped, "worst": wrong_fill[:6],
                           "missed_at": missed_at[:20], "extra_at": extra_at[:20]})
            mark = "OK" if (missed == 0 and extra == 0 and not wrong_fill) else "  "
            print("  %s %-40s %5d agreed  %4d missed  %4d extra  %4d fill off%s"
                  % (mark, source.stem[:40], agreed, missed, extra, len(wrong_fill),
                     "   (capped)" if capped else ""))
    finally:
        excel.Quit()

    caught = both + theirs_only
    clean = sum(1 for r in report
                if r["we_miss"] == 0 and r["we_add"] == 0 and r["fill_off"] == 0)
    print("\n%d of %d workbooks agree with Excel cell for cell" % (clean, len(report)))
    print("cells Excel formats: %d; we catch %d of them (%.4f), and %d Excel does not"
          % (caught, both, both / max(caught, 1), ours_only))
    print("of the agreed cells, %d wear the same fill (%.4f)"
          % (fill_same, fill_same / max(both, 1)))
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(
        {"agreed": both, "we_miss": theirs_only, "we_add": ours_only,
         "fill_same": fill_same, "docs": report}, ensure_ascii=False, indent=1),
        encoding="utf-8")
    print("written to %s" % args.out)
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
