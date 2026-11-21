"""What Excel's PivotTable object model answers, for the VBA host to match.

Builds a small source (Region, Product, Month, Sales) and a pivot with Region
and Product on rows, Month on columns and Kind as a page field, then reads the
members a macro uses -- the PivotTables collection, TableRange1/2, the field
collections and their orientation, position and function, the cache's source
and record count, Range.PivotTable -- and what RefreshTable does after the
source changes. Through COM rather than VBA, so the VBA project's trust
setting is never touched.

Usage: python tools/metrics/_xlsx_vba_pivot_probe.py
"""
import win32com.client

ROWS = [
    ("Region", "Product", "Month", "Kind", "Sales"),
    ("West", "Apple", "Feb", "A", 5),
    ("East", "Apple", "Jan", "B", 2),
    ("East", "Pear", "Feb", "A", 3),
    ("West", "Apple", "Jan", "A", 1),
]


def main():
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Add()
        d = wb.Worksheets(1)
        d.Name = "D"
        for i, row in enumerate(ROWS, 1):
            for j, v in enumerate(row, 1):
                d.Cells(i, j).Value = v
        p = wb.Worksheets.Add(After=d)
        p.Name = "P"
        pc = wb.PivotCaches().Create(1, "D!R1C1:R5C5")
        pt = pc.CreatePivotTable("P!R4C1", "PT")
        pt.PivotFields("Region").Orientation = 1
        pt.PivotFields("Product").Orientation = 1
        pt.PivotFields("Month").Orientation = 2
        pt.PivotFields("Kind").Orientation = 3
        pt.AddDataField(pt.PivotFields("Sales"))
        out = []
        say = lambda k, v: out.append(f"{k} = {v!r}")
        say("PivotTables.Count", p.PivotTables().Count)
        say("Name", pt.Name)
        say("TableRange1", pt.TableRange1.Address)
        say("TableRange2", pt.TableRange2.Address)
        say("PivotFields.Count", pt.PivotFields().Count)
        say("PivotFields", [f.Name for f in pt.PivotFields()])
        for coll in ("RowFields", "ColumnFields", "PageFields", "DataFields"):
            c = getattr(pt, coll)
            say(coll, [(f.Name, f.Orientation, f.Position) for f in c])
        df = pt.DataFields.Item(1)
        say("DataFields(1)", (df.Name, df.Caption, df.SourceName, df.Function, df.Orientation, df.Position, df.Value))
        say("Region", (pt.PivotFields("Region").Orientation, pt.PivotFields("Region").Position,
                       pt.PivotFields("Region").SourceName, pt.PivotFields("Region").Caption))
        say("Sales field", (pt.PivotFields("Sales").Orientation,))
        say("Kind CurrentPage", pt.PageFields.Item(1).CurrentPage.Name)
        say("SourceData", pt.SourceData)
        say("PivotCache.SourceData", pt.PivotCache().SourceData)
        say("RecordCount", pt.PivotCache().RecordCount)
        say("RowGrand/ColumnGrand", (pt.RowGrand, pt.ColumnGrand))
        say("Range(C7).PivotTable.Name", p.Range("C7").PivotTable.Name)
        say("PivotTables('PT').Name", p.PivotTables("PT").Name)
        say("value B7 before", p.Range("B7").Value)
        d.Range("E3").Value = 20
        say("value B7 unrefreshed", p.Range("B7").Value)
        say("RefreshTable", pt.RefreshTable())
        say("value B7 refreshed", p.Range("B7").Value)
        d.Range("E3").Value = 30
        say("PivotCache.Refresh", pt.PivotCache().Refresh())
        say("value B7 cache-refreshed", p.Range("B7").Value)
        d.Range("E3").Value = 40
        wb.RefreshAll()
        say("value B7 RefreshAll", p.Range("B7").Value)
        say("grid", [[p.Cells(r, c).Text for c in range(1, 5)] for r in range(1, 12)])
        print("\n".join(out))
        wb.Close(False)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()
