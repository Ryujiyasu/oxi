"""GETPIVOTDATA truth: Excel's own answers over the pivot render books.

Each book under pipeline_data/xlsx_pivot_render/books gets a sheet "Q" of
GETPIVOTDATA formulas -- grand totals by value-field caption and by source
field name, every outer item, outer x inner pairs, row x column pairs, and the
pairs Excel refuses (an inner field alone, an item that is not there, a hidden
item, a value field that is not there) -- and Excel calculates and saves them
into pipeline_data/xlsx_getpivotdata/. The saved values are the mark scheme:
`cargo run --example _corpus_formulas -- <books>` recalculates every formula
and reports the ones Oxi answers differently.

Usage (from the repository root): python tools/metrics/_xlsx_getpivotdata_gen.py
"""
import glob
import os

import win32com.client

# Run from the checkout whose pipeline_data holds the pivot books.
REPO = os.getcwd()
SRC = os.path.join(REPO, 'pipeline_data', 'xlsx_pivot_render', 'books')
OUT = os.path.join(REPO, 'pipeline_data', 'xlsx_getpivotdata')


def lit(v):
    """A formula literal for an item name."""
    try:
        float(v)
        return v
    except ValueError:
        return '"' + v.replace('"', '""') + '"'


def visible(field):
    return [it.Name for it in field.PivotItems() if it.Visible]


def main():
    os.makedirs(OUT, exist_ok=True)
    xl = win32com.client.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        for path in sorted(glob.glob(os.path.join(SRC, '*.xlsx'))):
            wb = xl.Workbooks.Open(path)
            ws = wb.Worksheets('Pivot')
            pt = ws.PivotTables(1)
            ref = "Pivot!" + pt.TableRange1.Cells(1, 1).Address
            datas = [d for d in pt.DataFields]
            # COM lists a table's fields in no fixed order; Position is theirs.
            rows = sorted((f for f in pt.RowFields if f.Name not in ('Values', '値')), key=lambda f: f.Position)
            cols = sorted((f for f in pt.ColumnFields if f.Name not in ('Values', '値')), key=lambda f: f.Position)
            formulas = []
            for d in datas:
                formulas.append(f'=GETPIVOTDATA({lit(d.Name)},{ref})')
                formulas.append(f'=GETPIVOTDATA({lit(d.SourceName)},{ref})')
            d0 = lit(datas[0].SourceName)
            for axis in (rows, cols):
                if not axis:
                    continue
                outer = axis[0]
                for item in visible(outer):
                    formulas.append(f'=GETPIVOTDATA({d0},{ref},{lit(outer.Name)},{lit(item)})')
                if len(axis) > 2:
                    # The innermost field with the outermost: one cell or many.
                    last = axis[-1]
                    for a in visible(outer)[:4]:
                        formulas.append(
                            f'=GETPIVOTDATA({d0},{ref},{lit(outer.Name)},{lit(a)},{lit(last.Name)},{lit(visible(last)[0])})')
                if len(axis) > 1:
                    inner = axis[1]
                    for a in visible(outer)[:2]:
                        for b in visible(inner)[:3]:
                            formulas.append(
                                f'=GETPIVOTDATA({d0},{ref},{lit(outer.Name)},{lit(a)},{lit(inner.Name)},{lit(b)})')
                    # An inner field without its outer one.
                    formulas.append(f'=GETPIVOTDATA({d0},{ref},{lit(inner.Name)},{lit(visible(inner)[0])})')
                # Case folded, and an item that is not there.
                first = visible(outer)[0]
                formulas.append(f'=GETPIVOTDATA({d0},{ref},{lit(outer.Name.upper())},{lit(first.lower())})')
                formulas.append(f'=GETPIVOTDATA({d0},{ref},{lit(outer.Name)},"NoSuchItem")')
                if any(it.Name == '(blank)' for it in outer.PivotItems()):
                    formulas.append(f'=GETPIVOTDATA({d0},{ref},{lit(outer.Name)},"(空白)")')
                hidden = [it.Name for it in outer.PivotItems() if not it.Visible]
                for h in hidden[:1]:
                    formulas.append(f'=GETPIVOTDATA({d0},{ref},{lit(outer.Name)},{lit(h)})')
            if rows and cols:
                for a in visible(rows[0])[:3]:
                    for b in visible(cols[0])[:3]:
                        formulas.append(
                            f'=GETPIVOTDATA({d0},{ref},{lit(rows[0].Name)},{lit(a)},{lit(cols[0].Name)},{lit(b)})')
            formulas.append(f'=GETPIVOTDATA("NoSuchValue",{ref})')
            q = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            q.Name = 'Q'
            for i, f in enumerate(formulas, start=1):
                q.Cells(i, 1).Value = "'" + f
                q.Cells(i, 2).Formula = f
            xl.CalculateFull()
            dst = os.path.join(OUT, os.path.basename(path))
            if os.path.exists(dst):
                os.remove(dst)
            wb.SaveAs(dst, 51)
            wb.Close(False)
            print(os.path.basename(path), len(formulas))
    finally:
        xl.Quit()


if __name__ == '__main__':
    main()
