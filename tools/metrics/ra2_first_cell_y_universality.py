"""§13.6 First Cell First-Char Y — universality sweep on bottom-15 docs.

c245888 commit had data: 8 tables across 3 docs (b35123, 2ea81a, 1ec1091)
all show Word dy=0 (first cell first-char y == table top y).

This expands the dataset by measuring 4 more table-heavy bottom-15 docs:
  - 04b88e (rank 8): 7 tables
  - 6514f (rank 12): 6 tables
  - d4d126 (rank 9): 5 tables
  - 459f (rank 11): 2 tables
Plus re-measure 2ea81a (rank 4): 4 tables, b35123 (rank 6): 2 tables, 1ec1 (rank 7): 1 table.

Goal: total 27+ tables across 7 docs all confirming dy=0.

For each table:
  table.Range.Information(6) = table_top_y
  first_cell.Range.Information(6) = first cell content top y
  dy = first_cell - table_top
"""
import os
import sys
import time
import json

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUT_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_first_cell_y_universality.json")

DOCX_DIR = "C:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"

DOCS_TO_MEASURE = [
    "04b88e7e0b25",
    "6514f214e482",
    "d4d126dfe1d9",
    "459f05f1e877",
    "2ea81a8441cc",
    "b35123fe8efc",
    "1ec1091177b1",
]


def find_docx(stem):
    for f in os.listdir(DOCX_DIR):
        if f.startswith(stem) and f.endswith(".docx"):
            return os.path.join(DOCX_DIR, f)
    return None


def measure_doc(word, path):
    """For each table, return (table_top_y, first_cell_y, dy)."""
    # Retry Open up to 3x with cleanup between attempts
    last_err = None
    for attempt in range(3):
        try:
            wdoc = word.Documents.Open(path)
            break
        except Exception as e:
            last_err = e
            time.sleep(2)
            try:
                # close any orphan docs
                while word.Documents.Count > 0:
                    word.Documents(1).Close(False)
            except Exception:
                pass
    else:
        raise last_err
    try:
        wdoc.Repaginate()
        time.sleep(0.1)
        results = []
        n_tables = wdoc.Tables.Count
        for ti in range(1, n_tables + 1):
            try:
                tbl = wdoc.Tables(ti)
                tbl_top = round(tbl.Range.Information(6), 4)
                first_cell = tbl.Cell(1, 1)
                fc_first_char = first_cell.Range.Characters(1)
                fc_y = round(fc_first_char.Information(6), 4)
                dy = round(fc_y - tbl_top, 4)
                # Try to get first paragraph y as alternative
                fc_para_y = round(first_cell.Range.Paragraphs(1).Range.Information(6), 4)
                dy_para = round(fc_para_y - tbl_top, 4)
                results.append({
                    "table_idx": ti,
                    "tbl_top_y": tbl_top,
                    "first_cell_first_char_y": fc_y,
                    "first_cell_first_para_y": fc_para_y,
                    "dy_char": dy,
                    "dy_para": dy_para,
                })
            except Exception as e:
                results.append({"table_idx": ti, "error": str(e)})
        return results
    finally:
        wdoc.Close(False)


def restart_word():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(3.0)
    return word


def main():
    word = restart_word()

    summary = {}
    try:
        for stem in DOCS_TO_MEASURE:
            path = find_docx(stem)
            if not path:
                print(f"  {stem}: no docx found")
                continue
            print(f"\n{stem}:")
            # Up to 2 word-level restarts per doc
            for attempt in range(2):
                try:
                    results = measure_doc(word, path)
                    summary[stem] = results
                    break
                except Exception as e:
                    msg = str(e)
                    if "RPC" in msg or "拒否" in msg or "コール" in msg:
                        print(f"  attempt {attempt+1}: COM failure → restart Word")
                        try:
                            word.Quit()
                        except Exception:
                            pass
                        time.sleep(5)
                        word = restart_word()
                        continue
                    else:
                        summary[stem] = {"error": msg}
                        print(f"  ERR: {msg}")
                        break
            else:
                summary[stem] = {"error": "RPC repeated failure after 2 word restarts"}
                print(f"  ERR: RPC repeated failure")
                continue
            try:
                results = summary[stem]
                for r in results:
                    if "error" in r:
                        print(f"  tbl{r['table_idx']}: ERR {r['error']}")
                    else:
                        print(f"  tbl{r['table_idx']}: tbl_top={r['tbl_top_y']:>7.2f}"
                              f" fc_char={r['first_cell_first_char_y']:>7.2f}"
                              f" dy_char={r['dy_char']:>+5.2f}"
                              f"  fc_para={r['first_cell_first_para_y']:>7.2f}"
                              f" dy_para={r['dy_para']:>+5.2f}")
            except Exception as e:
                print(f"  ERR: {e}")
                summary[stem] = {"error": str(e)}
    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)
        print(f"\nSaved to {OUT_JSON}")
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
