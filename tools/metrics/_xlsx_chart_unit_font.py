# -*- coding: utf-8 -*-
"""Does the tick spacing Excel allows follow the size of the axis's labels?

The sweep in `_xlsx_chart_unit.py` puts the fewest points Excel will leave
between two ticks at about 13.75pt with the default 10pt labels. Either that
is a fixed distance or it is the height of a label, and only a chart whose
labels are a different size can tell the two apart.

    python tools\\metrics\\_xlsx_chart_unit_font.py
"""
import sys

import win32com.client as com

sys.stdout.reconfigure(encoding="utf-8")

SIZES = [6, 8, 10, 14, 20, 28]
# One range per cap, so the chosen unit reports the cap it was chosen under.
RANGES = [(0, 100), (0, 350), (0, 1), (0, 12)]

excel = com.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
book = excel.Workbooks.Add()
sheet = book.Worksheets(1)
for row in range(1, 6):
    sheet.Cells(row, 1).Value = row * 10

print("size\tframe_pt\tinside_pt\tlow\thigh\tunit\tintervals")
try:
    for size in SIZES:
        for height in (180, 300):
            held = sheet.ChartObjects().Add(10, 10, 400, height)
            chart = held.Chart
            chart.SetSourceData(sheet.Range("A1:A5"))
            chart.ChartType = 4  # xlLine
            axis = chart.Axes(2)  # xlValue
            axis.TickLabels.Font.Size = size
            for low, high in RANGES:
                axis.MinimumScale = low
                axis.MaximumScale = high
                unit = axis.MajorUnit
                inside = chart.PlotArea.InsideHeight
                print(f"{size}\t{height}\t{inside:.2f}\t{low}\t{high}\t{unit}"
                      f"\t{(high - low) / unit:.4g}")
            held.Delete()
finally:
    book.Close(SaveChanges=False)
    excel.Quit()
