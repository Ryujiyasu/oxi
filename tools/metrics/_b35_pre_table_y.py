"""Find b35 first table top y in Oxi vs Word — is the PRE-table position correct?

If Oxi's table_top_y is already wrong before any cell-level code runs,
the §13.6 fix attempts (cell-level) couldn't possibly succeed.
"""
import os, sys, time, json
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import win32com.client

DOCX_REAL = "C:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
DOC = "b35123fe8efc_tokumei_08_01.docx"


def restart_word():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(3.0)
    return word


def main():
    word = restart_word()
    path = os.path.join(DOCX_REAL, DOC)
    last_err = None
    for attempt in range(3):
        try:
            wdoc = word.Documents.Open(path)
            break
        except Exception as e:
            last_err = e
            time.sleep(2)
            try:
                while word.Documents.Count > 0:
                    word.Documents(1).Close(False)
            except Exception:
                pass
    else:
        print(f"Failed: {last_err}")
        return
    try:
        wdoc.Repaginate()
        time.sleep(0.5)
        # Word's first table top y
        tbl = wdoc.Tables(1)
        tbl_top = round(tbl.Range.Information(6), 4)
        # First cell first char
        fc = tbl.Cell(1, 1).Range.Characters(1)
        fc_y = round(fc.Information(6), 4)
        # Word's first paragraph y (before any table)
        p1 = wdoc.Paragraphs(1)
        p1_y = round(p1.Range.Information(6), 4)
        # Last paragraph before first table
        # Walk paragraphs until we find one that's inside the table
        pre_tbl_p = None
        pre_tbl_y = None
        for i in range(1, min(20, wdoc.Paragraphs.Count + 1)):
            p = wdoc.Paragraphs(i)
            if p.Range.Tables.Count > 0:
                # This paragraph is in a table
                if pre_tbl_p is None:
                    pre_tbl_p = i - 1
                    if pre_tbl_p > 0:
                        pre_tbl_y = round(wdoc.Paragraphs(pre_tbl_p).Range.Information(6), 4)
                break
        print(f'Word b35 p.1:')
        print(f'  para 1 y: {p1_y}')
        if pre_tbl_p:
            print(f'  para {pre_tbl_p} (last before table): y={pre_tbl_y}')
        print(f'  table 1 top: {tbl_top}')
        print(f'  first cell first char: {fc_y}')
        print(f'  table-top - last_pre_table_y: {tbl_top - pre_tbl_y if pre_tbl_y else "N/A"}')
        print(f'  fc_y - tbl_top: {fc_y - tbl_top}')
    finally:
        wdoc.Close(False)
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
