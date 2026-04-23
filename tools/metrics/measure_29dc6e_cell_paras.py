"""Measure Y position of consecutive paragraphs within the
'代表者又は管理人の氏名' cell of 法人等 table in 29dc6e.

Previous attempt showed all 4 paragraphs at y=390.50. That's probably because
the cell contains paragraphs but COM was reporting only a subset. Here we
select each Paragraph index by range and pick line-level Y.

Expected identification:
- Para with text '代表者又は' → y = cell-top + offset
- Para with text '管理人の氏名（フリガナ）' → y = cell-top + offset + line_of_para1 + sa_p1 + sb_p2 (with or without collapse)
"""
import os
import json
import time
import win32com.client

DOC = os.path.abspath(
    "tools/golden-test/documents/docx/29dc6e8943fe_order_01.docx"
)
OUT = os.path.abspath(
    "tools/metrics/29dc6e_cell_paras_detail.json"
)


def main():
    app = win32com.client.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    rows = []
    try:
        doc = app.Documents.Open(DOC, ReadOnly=True)
        doc.Repaginate()
        time.sleep(1.0)

        # For each paragraph, also get selection info to measure exact Y of FIRST line
        paras = doc.Paragraphs
        for i in range(1, paras.Count + 1):
            rng = paras(i).Range
            # Select the range
            rng.Select()
            sel = app.Selection
            try:
                # Use Selection.Information to get more detail
                y_info = sel.Information(6)
                page = sel.Information(3)
                # HorizontalPositionRelativeToPage = 4 (the X)
                x_info = sel.Information(4)
            except Exception:
                y_info = -1
                page = -1
                x_info = -1
            text = rng.Text.rstrip("\r\n\x07")[:30]
            rows.append({
                "idx": i,
                "page": int(page),
                "y_pt": float(y_info),
                "x_pt": float(x_info),
                "text": text,
            })
            # Target area
            if 55 <= i <= 70:
                print(f"  idx={i:3d} p={page} y={y_info:6.2f} x={x_info:6.2f} text={text!r}")

        doc.Close(False)
    finally:
        app.Quit()

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"paragraphs": rows}, f, ensure_ascii=False, indent=2)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
