# -*- coding: utf-8 -*-
# Hold the renderer's per-row heights (OXI_XLSX_DUMP_ROWS) against Excel's
# (COM Rows(i).Height) for one workbook, and print the rows that disagree.
import os
import re
import subprocess
import sys
import win32com.client

RENDERER = r"tools\oxi-xlsx-renderer\target\release\oxi-xlsx-renderer.exe"


def main():
    path = os.path.abspath(sys.argv[1])
    out = os.environ.get("TEMP", ".") + r"\_row_model_probe.png"
    env = dict(os.environ, OXI_XLSX_DUMP_ROWS="1")
    run = subprocess.run([RENDERER, path, out], capture_output=True,
                         text=True, env=env)
    oxi = {}
    for line in run.stdout.splitlines():
        m = re.match(r"row (\d+) px (\S+)", line)
        if m:
            oxi[int(m.group(1))] = float(m.group(2))
    if not oxi:
        print("renderer dumped nothing:", run.stderr[:200])
        return

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(path, 0, True)
        ws = wb.Worksheets(1)
        n_diff = 0
        total_x = total_o = 0.0
        for r in sorted(oxi):
            x_px = ws.Rows(r).Height / 0.75
            o_px = oxi[r]
            total_x += x_px
            total_o += o_px
            if abs(x_px - o_px) > 1e-6:
                n_diff += 1
                if n_diff <= 30:
                    print("  row %-5d excel %6.1fpx oxi %6.1fpx  (%+.0f)" % (
                        r, x_px, o_px, o_px - x_px))
        print("rows differing: %d / %d   excel %.0fpx  oxi %.0fpx  (%+.0f)" % (
            n_diff, len(oxi), total_x, total_o, total_o - total_x))
        wb.Close(False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
