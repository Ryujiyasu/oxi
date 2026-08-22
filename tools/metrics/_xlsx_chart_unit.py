# -*- coding: utf-8 -*-
"""What Excel picks for a value axis's tick spacing when the chart does not say.

A chart part states `<c:max>` and leaves `<c:majorUnit>` out, so the spacing
between ticks — and so the number of labels down the axis — is Excel's own
decision. This asks Excel for it directly: build a line chart, pin the ends of
the axis and the height of the plot, and read back `MajorUnit`, which reports
the effective value whether or not it was set.

    python tools\\metrics\\_xlsx_chart_unit.py

Writes tab-separated rows: low, high, plot height in points, the unit Excel
chose, and how many intervals that makes.
"""
import sys

import win32com.client as com

sys.stdout.reconfigure(encoding="utf-8")

# Ends of the axis, and how tall the plot is. Both plausibly matter: a short
# plot has no room for many labels.
RANGES = [
    (0, 1), (0, 3), (0, 5), (0, 8), (0, 10), (0, 12), (0, 20), (0, 35),
    (0, 50), (0, 100), (0, 120), (0, 175), (0, 200), (0, 250), (0, 350),
    (0, 500), (0, 700), (0, 1000), (0, 2500), (0, 12000), (0, 0.35),
    (100, 350), (50, 100), (-50, 50), (0, 45), (0, 60), (0, 90),
]
HEIGHTS = [60, 90, 120, 150, 180, 240, 300, 380, 450, 600]

excel = com.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
book = excel.Workbooks.Add()
sheet = book.Worksheets(1)
for row in range(1, 6):
    sheet.Cells(row, 1).Value = row * 10

print("low\thigh\tframe_pt\tinside_pt\tunit\tintervals")
try:
    for height in HEIGHTS:
        held = sheet.ChartObjects().Add(10, 10, 400, height)
        chart = held.Chart
        chart.SetSourceData(sheet.Range("A1:A5"))
        chart.ChartType = 4  # xlLine
        axis = chart.Axes(2)  # xlValue
        for low, high in RANGES:
            axis.MinimumScale = low
            axis.MaximumScale = high
            unit = axis.MajorUnit
            # The plot the axis is drawn down, not the frame around it: the
            # category labels and the chart's margins take their share first.
            inside = chart.PlotArea.InsideHeight
            print(f"{low}\t{high}\t{height}\t{inside:.2f}\t{unit}"
                  f"\t{(high - low) / unit:.4g}")
        held.Delete()
finally:
    book.Close(SaveChanges=False)
    excel.Quit()
