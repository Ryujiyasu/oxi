# -*- coding: utf-8 -*-
"""Where Excel puts the ends of a value axis the chart does not pin.

A chart part that states no `<c:min>` leaves the foot of the axis to Excel,
and a graph drawn from 0 when Excel started at 200 is wrong everywhere. This
feeds Excel a series and reads back `MinimumScale` and `MaximumScale`.

    python tools\\metrics\\_xlsx_chart_ends.py
"""
import sys

import win32com.client as com

sys.stdout.reconfigure(encoding="utf-8")

# Each row is the numbers the chart plots.
SERIES = [
    [1, 2, 3],
    [10, 20, 30],
    [90, 130, 340],
    [300, 320, 340],
    [330, 335, 340],
    [0, 50, 100],
    [5, 5, 5],
    [-10, 0, 10],
    [-100, -50, -20],
    [0.1, 0.2, 0.35],
    [1000, 2000, 12000],
    [95, 96, 97],
    [50, 60, 70],
    [1, 100, 10000],
    [12, 12, 13],
    [-5, 20, 60],
    [200, 240, 260],
    [0, 0, 0],
]

excel = com.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
book = excel.Workbooks.Add()
sheet = book.Worksheets(1)

print("data\tlow\thigh\tunit\tinside_pt")
try:
    for numbers in SERIES:
        sheet.Range("A1:A20").Clear()
        for row, value in enumerate(numbers, start=1):
            sheet.Cells(row, 1).Value = value
        held = sheet.ChartObjects().Add(10, 10, 400, 300)
        chart = held.Chart
        chart.SetSourceData(sheet.Range(f"A1:A{len(numbers)}"))
        chart.ChartType = 4  # xlLine
        axis = chart.Axes(2)  # xlValue
        print(f"{numbers}\t{axis.MinimumScale}\t{axis.MaximumScale}"
              f"\t{axis.MajorUnit}\t{chart.PlotArea.InsideHeight:.2f}")
        held.Delete()
finally:
    book.Close(SaveChanges=False)
    excel.Quit()
