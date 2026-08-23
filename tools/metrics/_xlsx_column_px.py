# -*- coding: utf-8 -*-
r"""What width does Excel give each column, in pixels, past the drawn range?

`sanko_tool`'s panel hangs from column 5 to column 9, and the room its text is
broken in is the box between those two — but only columns 1 to 6 are in the
range the picture covers, so the last two are worked out from the sheet rather
than measured against Excel's own picture. A pixel there is the difference
between the panel's last line fitting and not.

Excel states a column's width itself: `Columns(n).Width` is points, which is
pixels at 96 dpi times three quarters. This asks Excel for every column of a
workbook and prints it beside what the renderer works out.

    python tools\metrics\_xlsx_column_px.py <workbook.xlsx> [columns]
"""
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"


def excel_widths(book, columns):
    """Each column's width in points and in pixels, from Excel itself."""
    script = f"""
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
  $wb = $excel.Workbooks.Open('{book}', 0, $true)
  $sh = $wb.Worksheets.Item(1)
  foreach ($n in 1..{columns}) {{
    $col = $sh.Columns.Item($n)
    Write-Output ("$n`t" + $col.Width + "`t" + $col.ColumnWidth)
  }}
  $wb.Close($false)
}} finally {{
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""
    done = subprocess.run(["powershell", "-NoProfile", "-Command", script],
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=600)
    held = {}
    for line in done.stdout.splitlines():
        parts = line.split("\t")
        if len(parts) == 3:
            try:
                held[int(parts[0])] = (float(parts[1]), float(parts[2]))
            except ValueError:
                continue
    if not held:
        print(done.stdout[-400:], done.stderr[-400:])
    return held


def ours(book):
    done = subprocess.run([str(RENDERER), str(book), str(Path(r"C:\tmp\_column_px.png")), "96"],
                          capture_output=True, timeout=600,
                          env={**__import__("os").environ, "OXI_XLSX_DUMP_COLUMNS": "1"})
    held = {}
    for line in done.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "column":
            held[int(parts[1])] = int(float(parts[3]))
    return held


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 2:
        print(__doc__)
        return
    book = Path(sys.argv[1]).resolve()
    columns = int(sys.argv[2]) if len(sys.argv) > 2 else 12
    told = excel_widths(book, columns)
    mine = ours(book)
    print(f"{'column':>7}{'Excel pt':>10}{'Excel px':>10}{'chars':>9}{'ours px':>9}")
    for number in sorted(told):
        points, chars = told[number]
        # A column's width in points is three quarters of its pixels.
        pixels = round(points * 4 / 3)
        # The dump counts columns from zero, and from the range's first.
        held = mine.get(number - 1)
        mark = "" if held is None or held == pixels else "  <<"
        print(f"{number:>7}{points:>10.2f}{pixels:>10}{chars:>9.2f}"
              f"{'-' if held is None else held:>9}{mark}")


if __name__ == "__main__":
    main()
