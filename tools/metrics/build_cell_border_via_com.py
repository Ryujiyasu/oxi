"""
Build cell-border-absorption fixtures via Word COM directly.

Sets cell.Borders(1) (left border) LineStyle + LineWidth via COM API,
bypassing python-docx XML which Word didn't recognize.

Sweep border widths {0, 0.25, 0.5, 0.75, 1.0, 1.5, 2.25, 3, 4.5, 6} pt.
"""
import os
import time
import win32com.client

OUT_DIR = os.path.join(os.path.dirname(__file__), "output",
                       "cell_border_absorption_v2")
os.makedirs(OUT_DIR, exist_ok=True)

# wdLineStyle enum
WD_LS_NONE = 0
WD_LS_SINGLE = 1

# wdLineWidth enum is in 1/4 pt units in older docs; modern is direct pt
# Modern: cell.Borders(1).LineWidth accepts wdLineWidth values:
#   wdLineWidth025pt = 2 (= 0.25pt)
#   wdLineWidth050pt = 4
#   wdLineWidth075pt = 6
#   wdLineWidth100pt = 8 (= 1.0pt)
#   wdLineWidth150pt = 12
#   wdLineWidth225pt = 18
#   wdLineWidth300pt = 24
#   wdLineWidth450pt = 36
#   wdLineWidth600pt = 48
LW_VALUES = {
    0.0:  None,  # use LineStyle=None
    0.25: 2,
    0.5:  4,
    0.75: 6,
    1.0:  8,
    1.5:  12,
    2.25: 18,
    3.0:  24,
    4.5:  36,
    6.0:  48,
}


def build_fixture(word, out_path, *, width_pt, left_padding=4.95):
    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72

    # Add a 2-row 1-col table
    rng = wdoc.Range(0, 0)
    tbl = wdoc.Tables.Add(rng, NumRows=2, NumColumns=1)

    # Set table-level cell padding (variable for sweep)
    tbl.LeftPadding = left_padding
    tbl.RightPadding = left_padding
    tbl.TopPadding = 0.0
    tbl.BottomPadding = 0.0

    # Set cell content
    for r in range(1, 3):
        cell = tbl.Cell(r, 1)
        cell.Range.Text = f"R{r}: text"
        cell.Range.Font.Name = "Calibri"
        cell.Range.Font.Size = 11

    # Apply LEFT border on row 1 cell only
    cell1 = tbl.Cell(1, 1)
    if width_pt == 0.0:
        cell1.Borders(1).LineStyle = WD_LS_NONE  # wdBorderLeft = 1
    else:
        lw = LW_VALUES[width_pt]
        cell1.Borders(1).LineStyle = WD_LS_SINGLE
        cell1.Borders(1).LineWidth = lw

    wdoc.SaveAs2(out_path)
    wdoc.Close(False)


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(2.0)
    # Sweep border width × padding
    pad_options = [0.0, 1.0, 2.5, 4.95]  # 4 padding levels
    border_options = [0.0, 0.25, 0.5, 0.75, 1.0, 1.5, 2.25, 3.0, 4.5, 6.0]
    try:
        for pad in pad_options:
            for w in border_options:
                pad_str = str(pad).replace(".", "p")
                w_str = str(w).replace(".", "p")
                name = f"CBV2_pad{pad_str}_w{w_str}pt.docx"
                path = os.path.join(OUT_DIR, name)
                try:
                    build_fixture(word, path, width_pt=w, left_padding=pad)
                    print(f"  built {name}")
                except Exception as e:
                    print(f"  ERR {name}: {e}")
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    print(f"\nFixtures saved to {OUT_DIR}")


if __name__ == "__main__":
    main()
