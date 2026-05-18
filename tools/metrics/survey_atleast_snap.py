"""S104 Phase B1 refinement: cross-context COM survey of atLeast row snap.

S98 vAlign sweep gave formula: visual = ceil((row_h + 0.5) / 0.75) × 0.75pt.
3 vAlign-sweep cases matched perfectly, but S103 SSIM verify showed
20 page regressions in real docs (b5f706 -0.0681, 34140b -0.0355, etc).

This script surveys EVERY atLeast row in all baseline docs and captures
Word's actual rendering vs the formula inputs:
  - trHeight value
  - content height estimate
  - rendered row pitch (y[next_row] - y[this_row])
  - table border width
  - font_size / family from first cell text
  - vAlign / cell count / paragraph count

Output: tools/metrics/atleast_snap_survey.json with rows like:
  {
    "doc": "459f05f1e877_kyodokenkyuyoushiki01.docx",
    "table_idx": 1, "row_idx": 3,
    "trH_pt": 11.4, "rendered_pitch_pt": 15.0,
    "border_pt": 0.5, "v_align": "top",
    "n_cells": 2, "n_paras_per_cell": 1,
    ...
  }

Analyze offline to derive context-aware formula.
"""
import json
import sys
from pathlib import Path
import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path('c:/Users/ryuji/oxi-main')
DOCX_DIR = ROOT / 'tools/golden-test/documents/docx'
OUT = ROOT / 'tools/metrics/atleast_snap_survey.json'

wdVerticalPositionRelativeToPage = 6
wdRowHeightAuto = 0
wdRowHeightAtLeast = 1
wdRowHeightExact = 2


def collapse_y(rng):
    doc = rng.Document
    return doc.Range(rng.Start, rng.Start).Information(wdVerticalPositionRelativeToPage)


def measure_doc(word, docx_path: Path) -> list[dict]:
    """For each atLeast row in each table, measure Word's actual pitch."""
    doc = word.Documents.Open(str(docx_path.absolute()), ReadOnly=True)
    rows_data = []
    try:
        for t_idx in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(t_idx)
            n_rows = tbl.Rows.Count
            try:
                border_pt = tbl.Borders.OutsideLineWidth or 0.5
            except Exception:
                border_pt = 0.5
            for r in range(1, n_rows + 1):
                row = tbl.Rows(r)
                try:
                    h_rule = row.HeightRule
                except Exception:
                    continue
                if h_rule == wdRowHeightExact:
                    continue  # only interested in atLeast / auto (= treated as atLeast)
                try:
                    trH_pt = row.Height  # in points
                except Exception:
                    trH_pt = None

                # Get this row's first cell first paragraph y
                try:
                    cell1 = tbl.Cell(Row=r, Column=1)
                    y_this = collapse_y(cell1.Range)
                except Exception:
                    continue

                # Get next row's first cell first paragraph y (= row pitch)
                pitch = None
                if r < n_rows:
                    try:
                        cell_next = tbl.Cell(Row=r + 1, Column=1)
                        y_next = collapse_y(cell_next.Range)
                        pitch = y_next - y_this
                    except Exception:
                        pass

                # Cell metadata: vAlign, paragraph count, first paragraph font
                try:
                    n_cells = row.Cells.Count
                except Exception:
                    n_cells = 1
                try:
                    v_align = cell1.VerticalAlignment  # 0=top 1=center 3=bottom
                except Exception:
                    v_align = None
                try:
                    n_paras = cell1.Range.Paragraphs.Count
                except Exception:
                    n_paras = 1
                try:
                    first_run = cell1.Range.Paragraphs(1).Range.Runs(1) if cell1.Range.Paragraphs(1).Range.Text.strip() else None
                    font_size = first_run.Font.Size if first_run else None
                    font_name = first_run.Font.NameFarEast if first_run else None
                except Exception:
                    font_size = None
                    font_name = None

                rows_data.append({
                    "doc": docx_path.name,
                    "table_idx": t_idx,
                    "row_idx": r - 1,  # 0-indexed for consistency
                    "trH_pt": trH_pt,
                    "rendered_pitch_pt": pitch,
                    "border_pt": border_pt,
                    "v_align": v_align,
                    "n_cells": n_cells,
                    "n_paras_per_cell": n_paras,
                    "font_size_pt": font_size,
                    "font_name": font_name,
                })
    finally:
        doc.Close(SaveChanges=False)
    return rows_data


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    all_rows = []
    try:
        docs = sorted([p for p in DOCX_DIR.glob('*.docx') if not p.name.startswith('test_')])
        print(f"Surveying {len(docs)} docs ...")
        for i, p in enumerate(docs):
            try:
                rows = measure_doc(word, p)
                all_rows.extend(rows)
                print(f"  [{i+1}/{len(docs)}] {p.name}: {len(rows)} atLeast rows")
            except Exception as e:
                print(f"  [{i+1}/{len(docs)}] {p.name}: ERROR {e}")
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(all_rows, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {len(all_rows)} row records to {OUT}")


if __name__ == '__main__':
    main()
