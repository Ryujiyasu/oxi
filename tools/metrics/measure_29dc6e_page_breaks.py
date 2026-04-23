"""Measure Word page attribution per paragraph for 29dc6e_order_01.

Purpose: identify where Oxi (7 pages) vs Word (6 pages) diverge.

Outputs JSON: { paragraph_index: {"page": N, "y_pt": Y, "text": "..." } }
"""
import os
import sys
import json
import time
import win32com.client

DOC = os.path.abspath(
    "tools/golden-test/documents/docx/29dc6e8943fe_order_01.docx"
)
OUT = os.path.abspath(
    "tools/metrics/29dc6e_word_para_pages.json"
)


def main():
    app = win32com.client.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        doc = app.Documents.Open(DOC, ReadOnly=True)
        # Let Word finish repagination
        doc.Repaginate()
        time.sleep(1.0)

        # wdVerticalPositionRelativeToPage = 6
        # wdActiveEndPageNumber = 3
        # wdFirstCharacterLineNumber = 10

        paras = doc.Paragraphs
        n = paras.Count
        print(f"paragraphs: {n}")
        rows = []
        prev_page = 0
        for i in range(1, n + 1):
            rng = paras(i).Range
            try:
                page = rng.Information(3)
                y = rng.Information(6)  # twips from top of page? actually pt from page top
            except Exception as e:
                page = -1
                y = -1.0
            text = rng.Text.rstrip("\r\n\x07")[:60]
            if page != prev_page:
                print(
                    f"  para[{i}] page={page} y={y:.2f} text={text!r}"
                )
                prev_page = page
            rows.append({"idx": i, "page": int(page), "y_pt": float(y), "text": text})

        doc.Close(False)
    finally:
        app.Quit()

    data = {
        "doc": os.path.basename(DOC),
        "paragraph_count": n,
        "paragraphs": rows,
    }
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
