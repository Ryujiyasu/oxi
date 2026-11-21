# -*- coding: utf-8 -*-
"""Build pivot-table truth workbooks with Excel itself.

One data sheet (Region / Product / Month / Sales / Qty, deterministic rows) and
one pivot per workbook, configured through COM (PivotCaches.Create ->
CreatePivotTable), refreshed and saved. Each workbook's pivot range is also
dumped as Excel shows it (Range.Formula / Value) to <name>.truth.json, which is
what an Oxi pivot computed from the same source must reproduce.

    python tools/metrics/_xlsx_pivot_gen.py
"""
import json, os, random
from pathlib import Path
import win32com.client

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / ("xlsx_pivot_r" if "--random" in __import__("sys").argv else "xlsx_pivot")
XL_SUM, XL_COUNT, XL_AVG, XL_MAX, XL_MIN = -4157, -4112, -4106, -4136, -4139
ROW, COL, PAGE, DATA = 1, 2, 3, 4

ARMS = {
    "p01_region_sum": dict(rows=["Region"], data=[("Sales", XL_SUM)]),
    "p02_region_product_sum": dict(rows=["Region", "Product"], data=[("Sales", XL_SUM)]),
    "p03_region_by_month": dict(rows=["Region"], cols=["Month"], data=[("Sales", XL_SUM)]),
    "p04_count": dict(rows=["Product"], data=[("Sales", XL_COUNT)]),
    "p05_average": dict(rows=["Region"], data=[("Qty", XL_AVG)]),
    "p06_two_data": dict(rows=["Region"], data=[("Sales", XL_SUM), ("Qty", XL_SUM)]),
    "p07_max_min": dict(rows=["Product"], data=[("Sales", XL_MAX), ("Sales", XL_MIN)]),
    "p08_page_filter": dict(rows=["Product"], pages=["Region"], data=[("Sales", XL_SUM)]),
    "p09_two_rows_cols": dict(rows=["Region", "Product"], cols=["Month"], data=[("Qty", XL_SUM)]),
    "p10_cols_only": dict(cols=["Product"], data=[("Sales", XL_SUM)]),
}
ARMS2 = {
    "q01_three_rows": dict(rows=["Region", "Product", "Month"], data=[("Sales", XL_SUM)]),
    "q02_two_cols": dict(rows=["Region"], cols=["Product", "Month"], data=[("Qty", XL_SUM)]),
    "q03_two_data_with_col": dict(rows=["Region"], cols=["Month"], data=[("Sales", XL_SUM), ("Qty", XL_AVG)]),
    "q04_numeric_rows": dict(rows=["Qty"], data=[("Sales", XL_COUNT)]),
    "q05_hidden_item": dict(rows=["Region"], data=[("Sales", XL_SUM)], hide=("Region", "North")),
    "q06_page_selected": dict(rows=["Product"], pages=["Region"], data=[("Sales", XL_SUM)], page_pick=("Region", "West")),
    "q07_blanks": dict(rows=["Region"], cols=["Month"], data=[("Sales", XL_SUM)], blanks=True),
    "q08_count_text": dict(rows=["Month"], data=[("Product", XL_COUNT)]),
    "q09_rows_and_page_two": dict(rows=["Month", "Region"], pages=["Product"], data=[("Qty", XL_MAX)]),
    "q10_data_on_rows": dict(rows=["Region"], data=[("Sales", XL_SUM), ("Qty", XL_SUM)], data_on_rows=True),
}


def data_rows():
    rnd = random.Random(20261121)
    regions = ["East", "West", "North", "South"]
    products = ["Apple", "Banana", "Cherry"]
    months = ["Jan", "Feb", "Mar"]
    out = []
    for _ in range(60):
        out.append([rnd.choice(regions), rnd.choice(products), rnd.choice(months),
                    rnd.randint(10, 500), rnd.randint(1, 20)])
    return out


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    rows = data_rows()
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        import sys
        arms = ARMS2 if "--second" in sys.argv else ARMS
        if "--random" in sys.argv:
            rnd = random.Random(int(sys.argv[sys.argv.index("--random") + 1]))
            arms = {}
            fields = ["Region", "Product", "Month"]
            for k in range(20):
                f = fields[:]
                rnd.shuffle(f)
                nrow = rnd.randint(0, 2)
                ncol = rnd.randint(0 if nrow else 1, 2 - min(nrow, 1) if nrow == 2 else 2)
                rows_, cols_ = f[:nrow], f[nrow:nrow + ncol]
                rest = f[nrow + ncol:]
                pages_ = rest[:1] if rest and rnd.random() < 0.4 else []
                nd = rnd.choice([1, 1, 2])
                data_ = [(rnd.choice(["Sales", "Qty"]), rnd.choice([XL_SUM, XL_COUNT, XL_AVG, XL_MAX, XL_MIN])) for _ in range(nd)]
                if nd == 2 and data_[0] == data_[1]:
                    data_ = data_[:1]
                arm = dict(rows=rows_, cols=cols_, pages=pages_, data=data_, blanks=rnd.random() < 0.3)
                if rows_ and rnd.random() < 0.3:
                    arm["hide"] = (rows_[0], rnd.choice({"Region": ["East", "West", "North", "South"], "Product": ["Apple", "Banana", "Cherry"], "Month": ["Jan", "Feb", "Mar"]}[rows_[0]]))
                arms[f"r{k:02d}"] = arm
        for name, arm in arms.items():
            wb = xl.Workbooks.Add()
            data = wb.Worksheets(1); data.Name = "Data"
            header = ["Region", "Product", "Month", "Sales", "Qty"]
            for j, h in enumerate(header, 1):
                data.Cells(1, j).Value = h
            for i, r in enumerate(rows, 2):
                for j, v in enumerate(r, 1):
                    if arm.get("blanks") and (i * 7 + j) % 11 == 0:
                        continue
                    data.Cells(i, j).Value = v
            src = data.Range(data.Cells(1, 1), data.Cells(len(rows) + 1, 5))
            pv = wb.Worksheets.Add(After=data); pv.Name = "Pivot"
            cache = wb.PivotCaches().Create(1, src)
            pt = cache.CreatePivotTable(pv.Range("A3"), "PT1")
            for f in arm.get("pages", []):
                pt.PivotFields(f).Orientation = PAGE
            for f in arm.get("rows", []):
                pt.PivotFields(f).Orientation = ROW
            for f in arm.get("cols", []):
                pt.PivotFields(f).Orientation = COL
            for f, fn in arm.get("data", []):
                pt.AddDataField(pt.PivotFields(f), None, fn)
            if arm.get("data_on_rows"):
                pt.DataPivotField.Orientation = ROW
            if arm.get("hide"):
                f, item = arm["hide"]
                pt.PivotFields(f).PivotItems(item).Visible = False
            if arm.get("page_pick"):
                f, item = arm["page_pick"]
                pt.PivotFields(f).CurrentPage = item
            pt.RefreshTable()
            used = pv.UsedRange
            grid = [[used.Cells(i, j).Value for j in range(1, used.Columns.Count + 1)]
                    for i in range(1, used.Rows.Count + 1)]
            truth = {"first_row": used.Row, "first_col": used.Column, "grid": grid,
                     "table_range": pt.TableRange2.Address}
            path = OUT / f"{name}.xlsx"
            if path.exists():
                path.unlink()
            wb.SaveAs(str(path), 51)
            wb.Close(False)
            (OUT / f"{name}.truth.json").write_text(json.dumps(truth, default=str, ensure_ascii=False, indent=0))
            print(name, truth["table_range"], len(grid), "x", len(grid[0]) if grid else 0)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()
