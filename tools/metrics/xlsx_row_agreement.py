# -*- coding: utf-8 -*-
"""How closely the renderer's row heights agree with Excel's, over a corpus.

SSIM answers "does the picture match" and mixes every cause together. This
answers one question sharply — is each row the height Excel gives it — which
is what the row-height model is judged on. Excel is opened once for the whole
run; the renderer is asked for its geometry per workbook.

    python tools/metrics/xlsx_row_agreement.py tools/golden-test/documents/xlsx
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


def oxi_rows(path: Path) -> dict:
    out = os.environ.get("TEMP", ".") + r"\_row_agreement.png"
    env = dict(os.environ, OXI_XLSX_DUMP_ROWS="1")
    run = subprocess.run([str(RENDERER), str(path), out], capture_output=True,
                         text=True, env=env)
    rows = {}
    for line in run.stdout.splitlines():
        m = re.match(r"row (\d+) px (\S+)", line)
        if m:
            rows[int(m.group(1))] = float(m.group(2))
    return rows


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("target", type=Path)
    parser.add_argument("--limit", type=int)
    parser.add_argument("--out", type=Path,
                        default=REPO / "pipeline_data" / "xlsx_row_agreement.json")
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
    total_rows = total_agree = 0
    try:
        for source in sources:
            rows = oxi_rows(source)
            if not rows:
                print("  %-40s renderer drew nothing" % source.stem[:40])
                continue
            try:
                wb = excel.Workbooks.Open(str(source.resolve()), 0, True)
            except Exception as error:
                print("  %-40s Excel would not open it: %s"
                      % (source.stem[:40], str(error)[:40]))
                continue
            ws = wb.Worksheets(1)
            agree = 0
            worst = []
            for index, ours in sorted(rows.items()):
                theirs = ws.Rows(index).Height / 0.75
                if abs(theirs - ours) < 1e-6:
                    agree += 1
                else:
                    worst.append((index, theirs, ours))
            wb.Close(False)
            total_rows += len(rows)
            total_agree += agree
            report.append({
                "doc": source.stem,
                "rows": len(rows),
                "agree": agree,
                "worst": worst[:8],
            })
            print("  %-44s %4d/%4d rows%s" % (
                source.stem[:44], agree, len(rows),
                "" if agree == len(rows) else "   first off: %s" % (worst[0],)))
    finally:
        excel.Quit()

    exact = sum(1 for r in report if r["agree"] == r["rows"])
    print("\n%d of %d workbooks match Excel row for row; %d of %d rows (%.4f)"
          % (exact, len(report), total_agree, total_rows,
             total_agree / max(total_rows, 1)))
    args.out.write_text(json.dumps(
        {"rows": total_rows, "agree": total_agree, "docs": report},
        ensure_ascii=False, indent=1), encoding="utf-8")
    print("written to %s" % args.out)
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
