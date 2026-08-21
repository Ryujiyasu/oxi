# -*- coding: utf-8 -*-
"""How closely the renderer's column widths agree with Excel's, over a corpus.

The sibling of `xlsx_row_agreement.py`. Rows are now exact on 281 of 285
workbooks; the width of a column is the other half of the geometry, and the
same question — is each column the width Excel gives it — is the sharp way to
ask it. Excel is opened once for the whole run.

    python tools/metrics/xlsx_column_agreement.py tools/golden-test/documents/xlsx
"""
import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"


def oxi_columns(path: Path) -> dict:
    out = os.environ.get("TEMP", ".") + r"\_column_agreement.png"
    env = dict(os.environ, OXI_XLSX_DUMP_COLUMNS="1")
    run = subprocess.run([str(RENDERER), str(path), out], capture_output=True,
                         text=True, env=env)
    columns = {}
    for line in (run.stdout or "").splitlines():
        m = re.match(r"column (\d+) px (\S+)", line)
        if m:
            columns[int(m.group(1))] = float(m.group(2))
    return columns


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("target", type=Path)
    parser.add_argument("--limit", type=int)
    parser.add_argument("--out", type=Path,
                        default=REPO / "pipeline_data" / "xlsx_column_agreement.json")
    args = parser.parse_args()

    sources = sorted(args.target.glob("*.xlsx")) if args.target.is_dir() else [args.target]
    sources = [p for p in sources if not p.name.startswith("~$")]
    if args.limit:
        sources = sources[: args.limit]

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    report = []
    total = agree = 0
    try:
        for source in sources:
            columns = oxi_columns(source)
            if not columns:
                print("  %-44s renderer drew nothing" % source.stem[:44])
                continue
            try:
                wb = excel.Workbooks.Open(str(source.resolve()), 0, True)
            except Exception as error:
                print("  %-44s Excel would not open it: %s"
                      % (source.stem[:44], str(error)[:40]))
                continue
            ws = wb.Worksheets(1)
            matched = 0
            worst = []
            for index, ours in sorted(columns.items()):
                # Columns are zero-based in the IR and one-based to Excel.
                theirs = ws.Columns(index + 1).Width / 0.75
                if ws.Columns(index + 1).Hidden:
                    theirs = 0.0
                if abs(theirs - ours) < 1e-6:
                    matched += 1
                else:
                    worst.append((index, theirs, ours))
            wb.Close(False)
            total += len(columns)
            agree += matched
            report.append({"doc": source.stem, "columns": len(columns),
                           "agree": matched, "worst": worst[:8]})
            print("  %-44s %4d/%4d columns%s" % (
                source.stem[:44], matched, len(columns),
                "" if matched == len(columns) else "   first off: %s" % (worst[0],)))
    finally:
        excel.Quit()

    exact = sum(1 for r in report if r["agree"] == r["columns"])
    print("\n%d of %d workbooks match Excel column for column; %d of %d columns (%.4f)"
          % (exact, len(report), agree, total, agree / max(total, 1)))
    args.out.write_text(json.dumps(
        {"columns": total, "agree": agree, "docs": report},
        ensure_ascii=False, indent=1), encoding="utf-8")
    print("written to %s" % args.out)
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
