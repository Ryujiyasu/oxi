# -*- coding: utf-8 -*-
"""How Excel draws compact-form drill buttons, read off its own pictures.

For each arm of `_xlsx_pivot_drill_gen.py` (font x size, buttons on/off):
the rows' pixel bands from COM (row tops relative to the used range), then
on the label column: the first ink x of a leaf label (Month, level 2) with
buttons off and on -- the shift D --, and on an outer row (Region, level 0;
Product, level 1) the button's box (its outline colour) and where the text
starts after it.

    python tools/metrics/_xlsx_pivot_drill_measure.py pipeline_data/xlsx_pivot_drill
"""
import glob
import json
import os
import sys
import time

import win32com.client
from PIL import Image


def retry(fn, tries=60):
    for k in range(tries):
        try:
            return fn()
        except Exception as bad:
            if "-2147418111" not in str(bad) or k == tries - 1:
                raise
            time.sleep(0.25)


def dark(p):
    return sum(p) < 330


def outline(p):
    r, g, b = p
    # the button's frame: a cool grey, never black text
    return 380 < r + g + b < 640 and b >= r and abs(r - g) < 30


def main():
    root = sys.argv[1]
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    out = {}
    try:
        for path in sorted(glob.glob(os.path.join(root, "*_d.xlsx"))):
            arm = os.path.basename(path)[:-7]
            pics = {}
            for v in "dn":
                pics[v] = Image.open(os.path.join(root, "diff", f"{arm}_{v}.excel.png")).convert("RGB")
            wb = retry(lambda: xl.Workbooks.Open(os.path.abspath(path), 0, True))
            ws = wb.Worksheets(1)
            used = ws.UsedRange
            top0 = ws.Rows(used.Row).Top
            left0 = ws.Columns(used.Column).Left
            a_w = round(ws.Columns(1).Width * 96 / 72)
            rows = {}
            for r in range(used.Row, used.Row + used.Rows.Count):
                text = ws.Cells(r, 1).Text
                y0 = round((ws.Rows(r).Top - top0) * 96 / 72)
                y1 = round((ws.Rows(r).Top + ws.Rows(r).Height - top0) * 96 / 72)
                rows.setdefault(text, (r, y0, y1))
            wb.Close(False)

            def ink_x(img, band, start=0, test=dark):
                _, y0, y1 = band
                for x in range(start, min(a_w, img.size[0])):
                    if any(test(img.getpixel((x, y))) for y in range(y0 + 1, y1 - 1)):
                        return x
                return None

            res = {"row_px": rows["East"][2] - rows["East"][1], "a_px": a_w}
            for label, key in (("Jan", "leaf"), ("East", "lv0"), ("Apple", "lv1")):
                band = rows.get(label)
                if not band:
                    continue
                n_x = ink_x(pics["n"], band)
                res[key + "_n"] = n_x
                if key == "leaf":
                    res["leaf_d"] = ink_x(pics["d"], band)
                    res["D"] = None if res["leaf_d"] is None or n_x is None else res["leaf_d"] - n_x
                    continue
                img = pics["d"]
                _, y0, y1 = band
                bx = [x for x in range(0, a_w) if any(outline(img.getpixel((x, y))) for y in range(y0 + 1, y1 - 1))]
                by = [y for y in range(y0, y1) if any(outline(img.getpixel((x, y))) for x in range(0, a_w))]
                if bx:
                    res[key + "_box"] = [min(bx), max(bx), min(by) - y0, max(by) - y0]
                    res[key + "_text_d"] = ink_x(img, band, max(bx) + 1)
            out[arm] = res
            print(arm, res)
    finally:
        xl.Quit()
    with open(os.path.join(root, "drill_measure.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
