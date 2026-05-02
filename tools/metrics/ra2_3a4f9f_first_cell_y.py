"""§13.6 verification on 3a4f9f — 108 tables.

If dy=0 across all 108, then §13.6 fix would address cumulative offset
on this doc. If 3a4f9f shows non-zero dy on some tables, that's a
new finding (unique to long doc).
"""
import os
import sys
import time
import json

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_JSON = os.path.join(os.path.dirname(__file__), "output",
                        "ra2_3a4f9f_first_cell_y.json")
DOCX_DIR = "C:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"


def find_3a4f9f():
    for f in os.listdir(DOCX_DIR):
        if f.startswith("3a4f9f") and f.endswith(".docx"):
            return os.path.join(DOCX_DIR, f)
    return None


def restart_word():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(3.0)
    return word


def main():
    path = find_3a4f9f()
    if not path:
        print("3a4f9f not found")
        return
    word = restart_word()
    results = []
    try:
        # Try opening with retry
        wdoc = None
        for attempt in range(3):
            try:
                wdoc = word.Documents.Open(path)
                break
            except Exception as e:
                print(f"Open attempt {attempt+1}: {e}")
                time.sleep(3)
                try:
                    while word.Documents.Count > 0:
                        word.Documents(1).Close(False)
                except Exception:
                    pass
        if wdoc is None:
            print("Failed to open after 3 attempts")
            return
        try:
            wdoc.Repaginate()
            time.sleep(0.5)
            n_tables = wdoc.Tables.Count
            print(f"3a4f9f tables: {n_tables}")
            zero_dy = 0
            nonzero = []
            for ti in range(1, n_tables + 1):
                try:
                    tbl = wdoc.Tables(ti)
                    tbl_top = round(tbl.Range.Information(6), 4)
                    fc = tbl.Cell(1, 1).Range.Characters(1)
                    fc_y = round(fc.Information(6), 4)
                    dy = round(fc_y - tbl_top, 4)
                    results.append({
                        "table_idx": ti,
                        "tbl_top_y": tbl_top, "fc_y": fc_y, "dy": dy,
                    })
                    if abs(dy) < 0.01:
                        zero_dy += 1
                    else:
                        nonzero.append((ti, dy))
                    if ti % 20 == 0:
                        print(f"  ... tbl {ti}/{n_tables}, dy_zero so far: {zero_dy}/{ti}")
                except Exception as e:
                    msg = str(e)[:60]
                    results.append({"table_idx": ti, "error": msg})
                    print(f"  tbl{ti}: ERR {msg}")
            print(f"\nFinal: zero_dy={zero_dy}/{n_tables}, nonzero={len(nonzero)}")
            if nonzero:
                print("Non-zero dy tables (potential carve-outs):")
                for ti, dy in nonzero[:20]:
                    print(f"  tbl{ti}: dy={dy:+.4f}pt")
        finally:
            wdoc.Close(False)
    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved to {OUT_JSON}")
        try: word.Quit()
        except: pass


if __name__ == "__main__":
    main()
