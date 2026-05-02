"""Get Word's first cell content text y (not just table border y)."""
import json
import sys
import time
from pathlib import Path
import pythoncom
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path("tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx")


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(str(DOCX.resolve()), ReadOnly=True)
        time.sleep(0.5)
        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(ti)
            tbl_top = tbl.Range.Information(6)
            print(f"\nTable {ti}: tbl.Range.Information(6) = {round(tbl_top,2)}")
            # Iterate cells via .Cells collection
            try:
                for ri in [1, 2, 3]:
                    for ci in [1, 2]:
                        try:
                            cell = tbl.Cell(ri, ci)
                        except Exception:
                            continue
                        cr = cell.Range
                        # Get first character position (skip cell marker)
                        # Cell range starts with text, ends with \x07 (cell marker)
                        try:
                            first_char_rng = doc.Range(cr.Start, cr.Start + 1)
                            first_y = first_char_rng.Information(6)
                            first_x = first_char_rng.Information(5)
                            text = (cr.Text or "")[:20].replace("\r","\\r").replace("\x07","\\x07")
                            print(f"  R{ri}C{ci}: first char ({first_x:.2f}, {first_y:.2f}) text='{text}'")
                        except Exception as e:
                            print(f"  R{ri}C{ci}: ERR {e}")
            except Exception as e:
                print(f"  Iteration ERR: {e}")
        doc.Close(False)
    finally:
        try: word.Quit()
        except: pass


if __name__ == "__main__":
    main()
