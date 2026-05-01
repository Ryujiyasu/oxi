"""Measure K_* 2ea81a-match repros."""
import json, sys
from pathlib import Path
import win32com.client as w32

REPRO_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\fe_match_repro")
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\fe_match_measurements.json")

# Reference: 2ea81a tbl#3 has Y0_intercept = 28.55pt (anchor=694.5, table_top=737.5, tblpY=14.45)
TBLPY_PT = 14.45  # all variants use tblpY=289tw


def main():
    docs = sorted(REPRO_DIR.glob("K_*.docx"))
    print(f"Measuring {len(docs)} docs...", file=sys.stderr)
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for d in docs:
            try:
                doc = word.Documents.Open(str(d.resolve()), ReadOnly=True)
            except Exception as e:
                results.append({"file": d.name, "error": f"open: {e}"})
                print(f"  ERROR open {d.name}: {e}", file=sys.stderr)
                continue
            try:
                tbl = doc.Tables(1)
                tt = tbl.Range.Information(6)
                tp = tbl.Range.Information(3)
                pre = doc.Range(0, tbl.Range.Start)
                last = pre.Paragraphs(pre.Paragraphs.Count)
                ay = last.Range.Information(6)
                paras = []
                for i in range(1, pre.Paragraphs.Count + 1):
                    p = pre.Paragraphs(i)
                    txt = (p.Range.Text or "")[:30].replace("\r","\\r").replace("\x07","\\x07")
                    paras.append({"idx": i, "y": p.Range.Information(6), "text": txt})
                results.append({
                    "file": d.name,
                    "anchor_y": ay,
                    "table_top": tt,
                    "table_page": tp,
                    "y0_intercept": tt - ay - TBLPY_PT,
                    "delta_anchor_to_table": tt - ay,
                    "paras": paras,
                })
                print(f"  done {d.name}", file=sys.stderr)
            except Exception as e:
                results.append({"file": d.name, "error": f"measure: {e}"})
                print(f"  ERROR measure {d.name}: {e}", file=sys.stderr)
            finally:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    print()
    print(f"{'file':30s} {'aY':>7} {'tY':>7} {'Δ':>7} {'Y0':>7}")
    print("-" * 65)
    for r in results:
        print(f"  {r['file']:28s}"
              f" {r['anchor_y']:>7.2f}"
              f" {r['table_top']:>7.2f}"
              f" {r['delta_anchor_to_table']:>+7.2f}"
              f" {r['y0_intercept']:>+7.2f}")
    print()
    print(f"  2ea81a tbl#3 reference Y0 intercept: +28.55pt")


if __name__ == "__main__":
    main()
