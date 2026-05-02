"""Measure 459f Word vs Oxi floating-table Y0.

Goal: identify the actual delta between Oxi's table_top and Word's
table_top for 459f's 2 floating tables, to ground any §19.10 fix.
"""
import os, sys, time, json
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import win32com.client

DOCX_REAL = "C:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
DOC = "459f05f1e877_kyodokenkyuyoushiki01.docx"


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False; word.DisplayAlerts = False; time.sleep(2.0)

    path = os.path.join(DOCX_REAL, DOC)
    last_err = None
    for attempt in range(3):
        try:
            wdoc = word.Documents.Open(path); break
        except Exception as e:
            last_err = e
            time.sleep(2)
            try:
                while word.Documents.Count > 0: word.Documents(1).Close(False)
            except: pass
    else:
        print(last_err); return

    try:
        wdoc.Repaginate(); time.sleep(0.5)

        for ti in range(1, wdoc.Tables.Count + 1):
            tbl = wdoc.Tables(ti)
            tbl_y = round(tbl.Range.Information(6), 4)
            try:
                page = tbl.Range.Information(3)  # wdActiveEndPageNumber
            except:
                page = "?"
            # Anchor paragraph
            try:
                anchor_p = tbl.Range.Paragraphs(1)
                anchor_y = round(anchor_p.Range.Information(6), 4)
            except:
                anchor_y = None
            # Find paragraph BEFORE the table
            tbl_start = tbl.Range.Start
            # Walk backwards to find non-table paragraph
            pre_y = None
            try:
                for i in range(min(50, wdoc.Paragraphs.Count + 1), 0, -1):
                    p = wdoc.Paragraphs(i)
                    if p.Range.End <= tbl_start and p.Range.Tables.Count == 0:
                        pre_y = round(p.Range.Information(6), 4)
                        break
            except: pass
            print(f'  Table {ti}: page={page} tbl_y={tbl_y} pre_table_para_y={pre_y}')
    finally:
        wdoc.Close(False)
        try: word.Quit()
        except: pass


if __name__ == "__main__":
    main()
