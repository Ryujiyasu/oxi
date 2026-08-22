# -*- coding: utf-8 -*-
"""How Excel lays out the entries of a legend inside the box it is given.

A chart that pins its legend with a manual layout states only the box; where
the entries go inside it is Excel's. The corpus's charts all carry the same
wide, short box and Excel puts the entries in a ROW across it, so only the
first is visible before the picture's edge — worth pinning properly rather
than guessing from one picture.

`LegendEntry` reports its own box in points, which is the measurement.

    python tools\\metrics\\_xlsx_chart_legend.py
"""
import sys

import win32com.client as com

sys.stdout.reconfigure(encoding="utf-8")

NAMES = ["男女計", "男", "女", "計", "そのほか"]
# (how many series, legend box width, height) as fractions of the chart.
SHAPES = [
    (3, 0.3786, 0.2197),
    (3, 0.3786, 0.5),
    (3, 0.15, 0.2197),
    (3, 0.15, 0.5),
    (2, 0.3786, 0.2197),
    (5, 0.3786, 0.2197),
    (5, 0.15, 0.6),
    (3, 0.9, 0.12),
]
FRAME = (400.0, 300.0)  # the chart's own box, in points
# The corpus's charts, to the point: three series named 男女計 / 男 / 女, a
# legend in ＭＳ 明朝 12pt, and the same box the template gives them.
LIKE_CORPUS = (3, 0.3786, 0.2197, "ＭＳ 明朝", 12.0, (763.0, 429.0))

excel = com.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
book = excel.Workbooks.Add()
sheet = book.Worksheets(1)

print("series\tbox_w\tbox_h\tentry\tleft\ttop\twidth\theight\tlegend_l\tlegend_t"
      "\tlegend_w\tlegend_h")
try:
    for count, wide, tall in SHAPES:
        sheet.Cells.Clear()
        for column in range(count):
            sheet.Cells(1, column + 1).Value = NAMES[column]
            for row in range(2, 6):
                sheet.Cells(row, column + 1).Value = row * 10 + column
        held = sheet.ChartObjects().Add(10, 10, FRAME[0], FRAME[1])
        chart = held.Chart
        chart.SetSourceData(sheet.Range(sheet.Cells(1, 1), sheet.Cells(5, count)))
        chart.ChartType = 4  # xlLine
        chart.HasLegend = True
        legend = chart.Legend
        legend.Left = 0.5369 * FRAME[0]
        legend.Top = 0.548 * FRAME[1]
        legend.Width = wide * FRAME[0]
        legend.Height = tall * FRAME[1]
        for index in range(1, count + 1):
            try:
                entry = legend.LegendEntries(index)
                print(f"{count}\t{wide}\t{tall}\t{index}\t{entry.Left:.2f}"
                      f"\t{entry.Top:.2f}\t{entry.Width:.2f}\t{entry.Height:.2f}"
                      f"\t{legend.Left:.2f}\t{legend.Top:.2f}\t{legend.Width:.2f}"
                      f"\t{legend.Height:.2f}")
            except Exception:
                # An entry Excel could not fit is not reported at all.
                print(f"{count}\t{wide}\t{tall}\t{index}\tdropped")
        held.Delete()

    # The same again at the corpus's own size and face.
    count, wide, tall, face, size, frame = LIKE_CORPUS
    sheet.Cells.Clear()
    for column in range(count):
        sheet.Cells(1, column + 1).Value = NAMES[column]
        for row in range(2, 6):
            sheet.Cells(row, column + 1).Value = row * 10 + column
    held = sheet.ChartObjects().Add(10, 10, frame[0], frame[1])
    chart = held.Chart
    chart.SetSourceData(sheet.Range(sheet.Cells(1, 1), sheet.Cells(5, count)))
    chart.ChartType = 4
    chart.HasLegend = True
    legend = chart.Legend
    legend.Font.Name = face
    legend.Font.Size = size
    legend.Left = 0.5369 * frame[0]
    legend.Top = 0.548 * frame[1]
    legend.Width = wide * frame[0]
    legend.Height = tall * frame[1]
    for index in range(1, count + 1):
        entry = legend.LegendEntries(index)
        print(f"corpus\t{wide}\t{tall}\t{index}\t{entry.Left:.2f}\t{entry.Top:.2f}"
              f"\t{entry.Width:.2f}\t{entry.Height:.2f}\t{legend.Left:.2f}"
              f"\t{legend.Top:.2f}\t{legend.Width:.2f}\t{legend.Height:.2f}")
    held.Delete()
finally:
    book.Close(SaveChanges=False)
    excel.Quit()
