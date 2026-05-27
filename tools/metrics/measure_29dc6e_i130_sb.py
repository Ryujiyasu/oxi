"""
S354 — COM-measure i=130 sb behavior in 29dc6e8943fe to verify S353 hypothesis.

Hypothesis: i=130 "（１）利用の区分" has w:before="146" (7.3pt sb). If Word
suppresses sb at first-paragraph-in-cell, then i=130's y = cell_top +
tcMar.top + 0 (no sb). If Word does NOT suppress, then i=130's y =
cell_top + tcMar.top + 7.3pt.

Oxi shows +6.45pt dy at i=131 (1 line below i=130). If Oxi adds extra
7.3pt sb that Word suppresses, the +6.45pt would be explained.

Measures:
1. The cell containing i=130 — RowIndex, ColIndex, Top.
2. Whether i=130 is the first paragraph in that cell.
3. i=130 y vs cell top (Word's actual position).
4. The previous paragraph (i=129) y vs i=130 y.
"""
import json
import sys
from pathlib import Path
import win32com.client

DOC_PATH = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\29dc6e8943fe_order_01.docx").resolve()
OUT_PATH = Path(r"c:\Users\ryuji\oxi-main\tools\metrics\29dc6e_i130_sb_word.json").resolve()

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = False

out = {"doc": str(DOC_PATH), "measurements": []}

try:
    doc = word.Documents.Open(str(DOC_PATH), ReadOnly=True)
    doc.ActiveWindow.View.Type = 3
    doc.Repaginate()

    # Locate paragraph i=130 by its known text "（１）利用の区分"
    needle = "（１）利用の区分"
    # Walk paragraphs to find by text
    target_para = None
    target_i = None
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        text = (p.Range.Text or "").rstrip("\r\x07")
        if needle in text:
            target_para = p
            target_i = i
            break
    if target_para is None:
        out["error"] = f"paragraph containing '{needle}' not found"
    else:
        out["i130_index"] = target_i
        rng = target_para.Range
        start = doc.Range(rng.Start, rng.Start)
        out["i130_y_pt"] = start.Information(6)
        out["i130_x_pt"] = start.Information(7)
        out["i130_page"] = start.Information(3)
        out["i130_text"] = (rng.Text or "").rstrip("\r\x07")[:60]

        # Get the cell (Information(12) = wdWithInTable; Information(13) = wdStartOfRangeRowNumber, etc.)
        # Use rng.Cells if it returns a cell (when range is within a cell)
        try:
            cells = rng.Cells
            if cells.Count > 0:
                cell = cells(1)
                out["cell_row_index"] = getattr(cell, "RowIndex", None)
                out["cell_col_index"] = getattr(cell, "ColumnIndex", None)
                out["cell_width_pt"] = cell.Width
                out["cell_height_pt"] = cell.Height
                # Cell top via collapsed-start range
                c_rng = cell.Range
                c_start = doc.Range(c_rng.Start, c_rng.Start)
                out["cell_start_y_pt"] = c_start.Information(6)
                out["cell_start_x_pt"] = c_start.Information(7)
                # Is i=130 the FIRST paragraph in the cell?
                first_para_in_cell = cell.Range.Paragraphs(1)
                out["cell_first_para_text"] = (first_para_in_cell.Range.Text or "").rstrip("\r\x07")[:60]
                out["i130_is_first_in_cell"] = (
                    first_para_in_cell.Range.Start == target_para.Range.Start
                )
                # Number of paragraphs in cell
                out["cell_n_paragraphs"] = cell.Range.Paragraphs.Count
                # First-para start y (= cell content top including padding)
                fp_rng = first_para_in_cell.Range
                fp_start = doc.Range(fp_rng.Start, fp_rng.Start)
                out["cell_first_para_y_pt"] = fp_start.Information(6)
        except Exception as e:
            out["cell_err"] = str(e)

        # Measure i=129 (previous paragraph)
        if target_i > 1:
            i_prev = doc.Paragraphs(target_i - 1)
            ip_rng = i_prev.Range
            ip_start = doc.Range(ip_rng.Start, ip_rng.Start)
            out["i129_y_pt"] = ip_start.Information(6)
            out["i129_x_pt"] = ip_start.Information(7)
            out["i129_text"] = (ip_rng.Text or "").rstrip("\r\x07")[:60]
            out["i129_to_i130_dy"] = out["i130_y_pt"] - out["i129_y_pt"]

        # Measure i=131 (next paragraph, the bullet "□ 研究")
        if target_i < doc.Paragraphs.Count:
            i_next = doc.Paragraphs(target_i + 1)
            in_rng = i_next.Range
            in_start = doc.Range(in_rng.Start, in_rng.Start)
            out["i131_y_pt"] = in_start.Information(6)
            out["i131_x_pt"] = in_start.Information(7)
            out["i131_text"] = (in_rng.Text or "").rstrip("\r\x07")[:60]
            out["i130_to_i131_dy"] = out["i131_y_pt"] - out["i130_y_pt"]

        # Read pPr sb from the paragraph format
        pf = target_para.Format
        out["i130_pf_SpaceBefore"] = pf.SpaceBefore  # points
        out["i130_pf_SpaceBeforeAuto"] = pf.SpaceBeforeAuto
        out["i130_pf_LineSpacing"] = pf.LineSpacing
        out["i130_pf_LineSpacingRule"] = pf.LineSpacingRule

    OUT_PATH.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote: {OUT_PATH}")
    print(json.dumps(out, ensure_ascii=False, indent=2))
finally:
    try:
        doc.Close(SaveChanges=False)
    except Exception:
        pass
    word.Quit()
