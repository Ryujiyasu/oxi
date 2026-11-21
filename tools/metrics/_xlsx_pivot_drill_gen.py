# -*- coding: utf-8 -*-
"""Compact pivots with and without drill buttons, across fonts and sizes.

From one Excel pivot book (q01_three_rows: Region > Product > Month on rows,
compact form) each arm sets the workbook's Normal font, CompactRowIndent 1
and ShowDrillIndicators on (d) or off (n), refreshes and saves. Excel's own
pictures (xlsx_pixel_diff.py) then show how far the buttons push the labels
and how big they are: `_xlsx_pivot_drill_measure.py` reads them.

    python tools/metrics/_xlsx_pivot_drill_gen.py <src.xlsx> <out-dir>
"""
import os
import sys
import time

import win32com.client

FONTS = [("cal", "Calibri"), ("ari", "Arial"), ("msp", "ＭＳ Ｐゴシック"), ("yu", "游ゴシック"), ("mei", "メイリオ")]
SIZES = [9, 10, 11, 12, 14, 16, 18, 20]
# OXI_DRILL_FONTS="tag:Face,tag:Face" and OXI_DRILL_SIZES="8,13,15" widen the sweep.
if os.environ.get("OXI_DRILL_FONTS"):
    FONTS = [tuple(item.split(":", 1)) for item in os.environ["OXI_DRILL_FONTS"].split(",")]
if os.environ.get("OXI_DRILL_SIZES"):
    SIZES = [int(v) for v in os.environ["OXI_DRILL_SIZES"].split(",")]
# CompactRowIndent per arm; OXI_DRILL_INDENTS="2,3" adds arms (drill on only)
# named <font><size>_i<n>_d.
INDENTS = [int(v) for v in os.environ.get("OXI_DRILL_INDENTS", "1").split(",")]


def retry(fn, tries=60):
    for k in range(tries):
        try:
            return fn()
        except Exception as bad:
            if "-2147418111" not in str(bad) or k == tries - 1:
                raise
            time.sleep(0.25)


def main():
    src, out = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
    os.makedirs(out, exist_ok=True)
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        for tag, face in FONTS:
            for size in SIZES:
                for indent, drill in [(i, d) for i in INDENTS for d in ((True, False) if i == 1 else (True,))]:
                    wb = retry(lambda: xl.Workbooks.Open(src))
                    retry(lambda: setattr(wb.Styles("Normal").Font, "Name", face))
                    retry(lambda: setattr(wb.Styles("Normal").Font, "Size", size))
                    ws = wb.Worksheets(1)
                    pt = ws.PivotTables(1)
                    retry(lambda: setattr(pt, "CompactRowIndent", indent))
                    retry(lambda: setattr(pt, "ShowDrillIndicators", drill))
                    retry(lambda: pt.RefreshTable())
                    dst = os.path.join(out, f"{tag}{size:02d}{'' if indent == 1 else f'_i{indent}'}_{'d' if drill else 'n'}.xlsx")
                    if os.path.exists(dst):
                        os.remove(dst)
                    retry(lambda: wb.SaveAs(dst, 51))
                    wb.Close(False)
    finally:
        xl.Quit()
    print(len(os.listdir(out)), "books")


if __name__ == "__main__":
    main()
