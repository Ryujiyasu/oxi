"""§13.6 First Cell Y rule — random baseline sample regression check.

The 7-doc/29-table sweep confirmed dy=0 in bottom-15 docs. To verify the
rule is universal (not bottom-bucket-specific), sample 20 random
baseline docs that have tables and measure dy.

Hypothesis: ALL baseline docs with tables show Word dy=0 (universal rule).
If any doc shows non-zero dy, that's a carve-out the spec needs to handle.
"""
import os
import sys
import time
import json
import random
import zipfile
import re

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUT_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_first_cell_y_baseline_sample.json")

DOCX_DIR = "C:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"

# Reproducible sample
random.seed(42)
SAMPLE_SIZE = 20

# Already-measured docs (skip)
ALREADY_MEASURED = {
    "04b88e7e0b25", "6514f214e482", "d4d126dfe1d9", "459f05f1e877",
    "2ea81a8441cc", "b35123fe8efc", "1ec1091177b1",
}


def docs_with_tables():
    """Return list of (stem, path, n_tables) for docs that contain tables."""
    found = []
    for f in sorted(os.listdir(DOCX_DIR)):
        if not f.endswith(".docx"):
            continue
        stem = f[:12]
        if stem in ALREADY_MEASURED:
            continue
        path = os.path.join(DOCX_DIR, f)
        try:
            with zipfile.ZipFile(path, "r") as zf:
                doc = zf.read("word/document.xml").decode("utf-8", errors="replace")
            n = len(re.findall(r'<w:tbl\b', doc))
            if n > 0:
                found.append((stem, path, n))
        except Exception:
            continue
    return found


def restart_word():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(3.0)
    return word


def measure_doc(word, path):
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
        raise last_err

    try:
        wdoc.Repaginate()
        time.sleep(0.1)
        results = []
        n_tables = wdoc.Tables.Count
        # Sample up to first 5 tables per doc to keep runtime reasonable
        max_tbls = min(n_tables, 5)
        for ti in range(1, max_tbls + 1):
            try:
                tbl = wdoc.Tables(ti)
                tbl_top = round(tbl.Range.Information(6), 4)
                first_cell = tbl.Cell(1, 1)
                fc_first_char = first_cell.Range.Characters(1)
                fc_y = round(fc_first_char.Information(6), 4)
                dy = round(fc_y - tbl_top, 4)
                results.append({
                    "table_idx": ti,
                    "tbl_top_y": tbl_top,
                    "first_cell_first_char_y": fc_y,
                    "dy_char": dy,
                })
            except Exception as e:
                results.append({"table_idx": ti, "error": str(e)[:80]})
        return {"n_tables": n_tables, "measured_tables": max_tbls,
                "tables": results}
    finally:
        wdoc.Close(False)


def main():
    candidates = docs_with_tables()
    print(f"Docs with tables (excluding already measured): {len(candidates)}")

    sample = random.sample(candidates, min(SAMPLE_SIZE, len(candidates)))
    sample.sort()
    print(f"Sampled {len(sample)} docs.")

    word = restart_word()
    summary = {}
    nonzero_dy = []

    try:
        for stem, path, n_tables_xml in sample:
            print(f"\n{stem} (xml n_tbls={n_tables_xml}):")
            for attempt in range(2):
                try:
                    res = measure_doc(word, path)
                    summary[stem] = res
                    for t in res["tables"]:
                        if "error" in t:
                            print(f"  tbl{t['table_idx']}: ERR {t['error']}")
                        else:
                            mark = "✓" if abs(t['dy_char']) < 0.01 else "✗"
                            print(f"  tbl{t['table_idx']}: dy={t['dy_char']:+5.2f} {mark}")
                            if abs(t['dy_char']) > 0.01:
                                nonzero_dy.append((stem, t))
                    break
                except Exception as e:
                    msg = str(e)
                    if "RPC" in msg or "拒否" in msg or "コール" in msg:
                        print(f"  attempt {attempt+1}: COM failure → restart Word")
                        try: word.Quit()
                        except: pass
                        time.sleep(5)
                        word = restart_word()
                        continue
                    summary[stem] = {"error": msg}
                    print(f"  ERR: {msg[:80]}")
                    break
            else:
                summary[stem] = {"error": "RPC repeated failure"}
    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)
        # Aggregate stats
        total_tables = 0
        zero_dy = 0
        for stem, res in summary.items():
            if isinstance(res, dict) and "tables" in res:
                for t in res["tables"]:
                    if "dy_char" in t:
                        total_tables += 1
                        if abs(t["dy_char"]) < 0.01:
                            zero_dy += 1
        print(f"\n=== Summary ===")
        print(f"Total tables measured: {total_tables}")
        print(f"dy == 0.00pt: {zero_dy} / {total_tables}")
        if nonzero_dy:
            print(f"\nNon-zero dy cases (carve-outs to investigate):")
            for stem, t in nonzero_dy:
                print(f"  {stem} tbl{t['table_idx']}: dy={t['dy_char']:+.2f}pt")
        print(f"\nSaved to {OUT_JSON}")
        try: word.Quit()
        except: pass


if __name__ == "__main__":
    main()
