# -*- coding: utf-8 -*-
"""Dump Excel's built-in pivot table styles element by element, as
_xlsx_table_styles_dump.py does for table styles. Pivot styles have their own
elements (subtotal rows and columns, row subheadings, page field labels, the
grand total row and column ...). Output: pipeline_data/xlsx_pivot_styles.json

    python tools/metrics/_xlsx_pivot_styles_dump.py
"""
import json, sys
from pathlib import Path

import win32com.client

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _xlsx_table_styles_dump import element  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "xlsx_pivot_styles.json"
# XlTableStyleElementType values a pivot style uses.
ELEMENTS = {0: "wholeTable", 1: "headerRow", 2: "grandTotalRow", 3: "firstColumn", 4: "grandTotalColumn",
            5: "firstRowStripe", 6: "secondRowStripe", 7: "firstColumnStripe", 8: "secondColumnStripe",
            9: "firstHeaderCell", 13: "subtotalColumn1", 14: "subtotalColumn2", 15: "subtotalColumn3",
            16: "subtotalRow1", 17: "subtotalRow2", 18: "subtotalRow3", 19: "blankRow",
            20: "columnSubheading1", 21: "columnSubheading2", 22: "columnSubheading3",
            23: "rowSubheading1", 24: "rowSubheading2", 25: "rowSubheading3",
            26: "pageFieldLabels", 27: "pageFieldValues"}


def main():
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    wb = xl.Workbooks.Add()
    styles = {}
    try:
        for fam, n in (("Light", 28), ("Medium", 28), ("Dark", 28)):
            for k in range(1, n + 1):
                name = f"PivotStyle{fam}{k}"
                try:
                    ts = wb.TableStyles(name)
                except Exception:
                    continue
                rec = {}
                for idx, el in ELEMENTS.items():
                    try:
                        got = element(ts.TableStyleElements(idx))
                    except Exception:
                        got = None
                    if got is not None:
                        rec[el] = got
                styles[name] = rec
        OUT.write_text(json.dumps(styles, indent=1, ensure_ascii=False), encoding="utf-8")
        print(f"wrote {len(styles)} pivot styles to {OUT}")
    finally:
        wb.Close(False); xl.Quit()


if __name__ == "__main__":
    main()
