"""Debug Tables 5+6 of a1d6 — Table 5 enumerated 0 cells but has 30 rows."""
from __future__ import annotations
import os, sys, traceback
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOC = os.path.abspath('tools/golden-test/documents/docx/a1d6e4efa2e7_tokumei_08_01-4.docx')


def info5(d, pos):
    try: return float(d.Range(pos, pos).Information(5))
    except: return None
def info6(d, pos):
    try: return float(d.Range(pos, pos).Information(6))
    except: return None


def main():
    word = wc.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.ScreenUpdating = False
    doc = None
    try:
        doc = word.Documents.Open(DOC, ReadOnly=True)
        doc.Repaginate()
        for ti in (5, 6):
            t = doc.Tables(ti)
            print(f'\n=== Table {ti}: rows={t.Rows.Count}, columns={getattr(t,"Columns",None) and t.Columns.Count} ===')
            for ri in range(1, min(t.Rows.Count, 35) + 1):
                try:
                    row = t.Rows(ri)
                    n_cells = row.Cells.Count
                    rtop = info6(doc, row.Range.Start)
                    print(f'  Row {ri}: n_cells={n_cells}, top_y={rtop}, height_rule={row.HeightRule}, height={row.Height}')
                    for ci in range(1, n_cells + 1):
                        try:
                            cell = row.Cells(ci)
                            n_p = cell.Range.Paragraphs.Count
                            first_p = cell.Range.Paragraphs(1)
                            text = (first_p.Range.Text or '').rstrip('\r\n\x07')[:30]
                            p_y = info6(doc, first_p.Range.Start)
                            sb = first_p.Format.SpaceBefore
                            style = first_p.Style.NameLocal if first_p.Style else None
                            print(f'    Cell {ci}: n_paras={n_p}, p1_y={p_y}, sb={sb:.2f}, style={style!r}, text={text!r}')
                        except Exception as e:
                            print(f'    Cell {ci}: error {e}')
                except Exception as e:
                    print(f'  Row {ri}: error {e}')
    except Exception:
        traceback.print_exc()
    finally:
        if doc is not None:
            doc.Close(SaveChanges=0)
        word.Quit()


if __name__ == '__main__':
    main()
