"""What Excel answers for a pivot field's items and for PivotTable.GetPivotData.

The same table as _xlsx_vba_pivot_probe.py (Region and Product on rows, Month
on columns, Kind as a page field, the sum of Sales), then: each field's
PivotItems (count, names in order, Value, Caption, Visible, Position,
RecordCount, LabelRange, DataRange), hiding an item and showing it again
(the grid each time), and GetPivotData's cell for several argument lists,
including ones that name no cell. Through COM, not VBA.

Usage: python tools/metrics/_xlsx_vba_pivot_items_probe.py
"""
import time

import pywintypes
import win32com.client

ROWS = [
    ("Region", "Product", "Month", "Kind", "Sales"),
    ("West", "Apple", "Feb", "A", 5),
    ("East", "Apple", "Jan", "B", 2),
    ("East", "Pear", "Feb", "A", 3),
    ("West", "Apple", "Jan", "A", 1),
]


def retry(fn):
    for _ in range(40):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.args[0] in (-2147418111, -2146777998):  # call rejected / busy
                time.sleep(0.25)
                continue
            raise


def attempt(fn):
    try:
        return retry(fn)
    except pywintypes.com_error as e:
        exc = e.args[2] if len(e.args) > 2 and e.args[2] else None
        return ("ERR", e.args[0], exc[2] if exc else None, exc[5] if exc else None)


def main():
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = retry(lambda: xl.Workbooks.Add())
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
        grid = lambda: [[p.Cells(r, c).Text for c in range(1, 6)] for r in range(1, 12)]
        say("grid", grid())
        for name in ("Region", "Product", "Month", "Kind", "Sales"):
            pf = pt.PivotFields(name)
            items = attempt(lambda: pf.PivotItems())
            if isinstance(items, tuple):
                say(f"{name}.PivotItems", items)
                continue
            say(f"{name}.PivotItems.Count", items.Count)
            for i in range(1, items.Count + 1):
                it = items.Item(i)
                say(f"{name}[{i}]", (
                    attempt(lambda: it.Name), attempt(lambda: it.Value), attempt(lambda: it.Caption),
                    attempt(lambda: it.Visible), attempt(lambda: it.Position),
                    attempt(lambda: it.RecordCount),
                    attempt(lambda: it.LabelRange.Address), attempt(lambda: it.DataRange.Address),
                    attempt(lambda: it.Parent.Name),
                ))
        say("Region.PivotItems('west').Name", attempt(lambda: pt.PivotFields("Region").PivotItems("west").Name))
        say("Region.PivotItems('Nope')", attempt(lambda: pt.PivotFields("Region").PivotItems("Nope").Name))
        say("Region.PivotItems(9)", attempt(lambda: pt.PivotFields("Region").PivotItems(9).Name))
        say("Sales data PivotItems", attempt(lambda: pt.DataFields.Item(1).PivotItems().Count))
        # GetPivotData
        cases = [
            ("Sales",),
            ("Sales", "Region", "West"),
            ("Sales", "Region", "West", "Product", "Apple"),
            ("Sales", "Month", "Jan"),
            ("Sales", "Region", "East", "Month", "Feb"),
            ("Sales", "Region", "East", "Product", "Pear", "Month", "Feb"),
            ("Sales", "Region", "East", "Product", "Pear", "Month", "Jan"),
            ("Sales", "Product", "Apple"),
            ("Sales", "Region", "Nope"),
            ("合計 / Sales", "Region", "West"),
            ("sales", "region", "west"),
            ("Sales", "Kind", "A"),
            (),
        ]
        for args in cases:
            r = attempt(lambda: pt.GetPivotData(*args))
            if isinstance(r, tuple):
                say(f"GetPivotData{args}", r)
            else:
                say(f"GetPivotData{args}", (r.Address, r.Value))
        # Hiding an item
        say("Visible=False East", attempt(lambda: setattr(pt.PivotFields("Region").PivotItems("East"), "Visible", False)))
        say("grid hidden East", grid())
        say("East.Visible", attempt(lambda: pt.PivotFields("Region").PivotItems("East").Visible))
        say("East.Position hidden", attempt(lambda: pt.PivotFields("Region").PivotItems("East").Position))
        say("Region.VisibleItems", attempt(lambda: [i.Name for i in pt.PivotFields("Region").VisibleItems]))
        say("Region.HiddenItems", attempt(lambda: [i.Name for i in pt.PivotFields("Region").HiddenItems]))
        say("Region count hidden", attempt(lambda: pt.PivotFields("Region").PivotItems().Count))
        say("GetPivotData East hidden", attempt(lambda: pt.GetPivotData("Sales", "Region", "East").Address))
        say("Visible=False West too", attempt(lambda: setattr(pt.PivotFields("Region").PivotItems("West"), "Visible", False)))
        say("Visible=True East", attempt(lambda: setattr(pt.PivotFields("Region").PivotItems("East"), "Visible", True)))
        say("grid shown", grid())
        R = lambda: pt.PivotFields("Region")
        attempt(lambda: setattr(R().PivotItems("East"), "Visible", False))
        say("West.Position, East hidden", attempt(lambda: R().PivotItems("West").Position))
        say("East.RecordCount hidden", attempt(lambda: R().PivotItems("East").RecordCount))
        say("East.LabelRange hidden", attempt(lambda: R().PivotItems("East").LabelRange.Address))
        say("Region item names hidden", attempt(lambda: [i.Name for i in R().PivotItems()]))
        attempt(lambda: setattr(R().PivotItems("East"), "Visible", True))
        say("Pear.Position=1", attempt(lambda: setattr(pt.PivotFields("Product").PivotItems("Pear"), "Position", 1)))
        say("grid Pear 1st", grid())
        say("Product names", attempt(lambda: [(i.Name, i.Position) for i in pt.PivotFields("Product").PivotItems()]))
        say("Pear.Position=5", attempt(lambda: setattr(pt.PivotFields("Product").PivotItems("Pear"), "Position", 5)))
        say("Pear.Position=2", attempt(lambda: setattr(pt.PivotFields("Product").PivotItems("Pear"), "Position", 2)))
        say("grid Pear 2nd", grid())
        say("Kind.CurrentPage=A", attempt(lambda: setattr(pt.PivotFields("Kind"), "CurrentPage", "A")))
        say("Kind items picked", attempt(lambda: [(i.Name, i.Visible) for i in pt.PivotFields("Kind").PivotItems()]))
        say("Kind A DataRange picked", attempt(lambda: pt.PivotFields("Kind").PivotItems("A").DataRange.Address))
        say("GetPivotData Kind A picked", attempt(lambda: pt.GetPivotData("Sales", "Kind", "A").Address))
        say("Kind.CurrentPage=(All)", attempt(lambda: setattr(pt.PivotFields("Kind"), "CurrentPage", "(All)")))
        say("Kind items all", attempt(lambda: [(i.Name, i.Visible) for i in pt.PivotFields("Kind").PivotItems()]))
        # Hidden page item: Kind B hidden (not a page choice)
        say("Kind B hidden", attempt(lambda: setattr(pt.PivotFields("Kind").PivotItems("B"), "Visible", False)))
        say("grid Kind B hidden", grid())
        say("Kind.CurrentPage", attempt(lambda: pt.PivotFields("Kind").CurrentPage.Name))
        say("Kind.EnableMultiplePageItems", attempt(lambda: pt.PivotFields("Kind").EnableMultiplePageItems))
        # Position change of an item
        say("Apple.Position=2", attempt(lambda: setattr(pt.PivotFields("Product").PivotItems("Apple"), "Position", 2)))
        say("grid Apple 2nd", grid())
        # Renaming an item caption
        say("Jan.Caption='January'", attempt(lambda: setattr(pt.PivotFields("Month").PivotItems("Jan"), "Caption", "January")))
        say("grid January", grid())
        say("Jan item name now", attempt(lambda: [i.Name for i in pt.PivotFields("Month").PivotItems()]))
        say("Jan item SourceName", attempt(lambda: pt.PivotFields("Month").PivotItems(1).SourceName))
        print("\n".join(out))
        wb.Close(False)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()
