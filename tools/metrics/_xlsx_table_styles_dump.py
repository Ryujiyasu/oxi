# -*- coding: utf-8 -*-
"""Dump Excel's built-in table styles, element by element, as Excel holds them.

Each element (whole table, header row, row stripes, first column, ...) is read
through COM: whether it carries a format, its fill, font and the six borders,
each colour as a theme slot + TintAndShade where Excel holds it that way, else
as RGB. Output: pipeline_data/xlsx_table_styles.json

    python tools/metrics/_xlsx_table_styles_dump.py
"""
import json
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "xlsx_table_styles.json"
ELEMENTS = {0: "wholeTable", 1: "headerRow", 2: "totalRow", 3: "firstColumn", 4: "lastColumn",
            5: "firstRowStripe", 6: "secondRowStripe", 7: "firstColumnStripe", 8: "secondColumnStripe",
            9: "firstHeaderCell", 10: "lastHeaderCell", 11: "firstTotalCell", 12: "lastTotalCell"}
BORDERS = {7: "left", 8: "top", 9: "bottom", 10: "right", 11: "insideV", 12: "insideH"}


def colour(obj):
    out = {}
    try:
        t = obj.ThemeColor
        if t:
            out["theme"] = int(t)
            out["tint"] = round(float(obj.TintAndShade), 6)
    except Exception:
        pass
    try:
        c = int(obj.Color)
        out["rgb"] = "%02X%02X%02X" % (c & 255, (c >> 8) & 255, (c >> 16) & 255)
    except Exception:
        pass
    return out


def element(e):
    if not e.HasFormat:
        return None
    rec = {}
    i = e.Interior
    try:
        if i.Pattern not in (-4142,):  # xlNone
            rec["fill"] = colour(i) | {"pattern": int(i.Pattern)}
    except Exception:
        pass
    f = e.Font
    font = {}
    try:
        if f.Bold is True:
            font["bold"] = True
    except Exception:
        pass
    try:
        if f.Italic is True:
            font["italic"] = True
    except Exception:
        pass
    c = colour(f)
    if c:
        font["colour"] = c
    if font:
        rec["font"] = font
    borders = {}
    for k, name in BORDERS.items():
        try:
            b = e.Borders(k)
            if b.LineStyle not in (-4142, None):
                borders[name] = {"style": int(b.LineStyle), "weight": int(b.Weight)} | colour(b)
        except Exception:
            pass
    if borders:
        rec["borders"] = borders
    return rec


def main():
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Add()
    styles = {}
    try:
        for fam, n in (("Light", 21), ("Medium", 28), ("Dark", 11)):
            for k in range(1, n + 1):
                name = f"TableStyle{fam}{k}"
                ts = wb.TableStyles(name)
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
        print(f"wrote {len(styles)} styles to {OUT}")
    finally:
        wb.Close(False)
        xl.Quit()


if __name__ == "__main__":
    main()
