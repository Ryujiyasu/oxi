"""Measure O_* order-test variants."""
import json, re, zipfile, sys
from pathlib import Path
import win32com.client as w32

REPRO_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\order_test_repro")
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\order_test_measurements.json")


def main():
    docs = sorted(REPRO_DIR.glob("O_*.docx"))
    print(f"Measuring {len(docs)} docs...", file=sys.stderr)
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for d in docs:
            with zipfile.ZipFile(d) as z:
                xml = z.read("word/document.xml").decode("utf-8")
            tw = int(re.search(r'w:tblpY="(-?\d+)"', xml).group(1))
            doc = word.Documents.Open(str(d.resolve()), ReadOnly=True)
            try:
                tbl = doc.Tables(1)
                tt = tbl.Range.Information(6)
                tp = tbl.Range.Information(3)
                pre = doc.Range(0, tbl.Range.Start)
                last = pre.Paragraphs(pre.Paragraphs.Count)
                ay = last.Range.Information(6)
                results.append({
                    "file": d.name,
                    "tblpY_pt": tw / 20.0,
                    "anchor_top_pt": ay,
                    "table_top_pt": tt,
                    "table_page": tp,
                    "delta_pt": tt - ay if (tt and ay) else None,
                })
            finally:
                doc.Close(SaveChanges=0)
            print(f"  done {d.name}", file=sys.stderr)
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)

    print()
    print(f"{'file':30s} {'tblpY_pt':>9} {'a_top':>7} {'t_top':>7} {'delta':>7}")
    print("-" * 70)
    for r in results:
        print(f"  {r['file']:28s}"
              f" {r['tblpY_pt']:>9.2f}"
              f" {r['anchor_top_pt']:>7.2f}"
              f" {r['table_top_pt']:>7.2f}"
              f" {r['delta_pt']:>+7.2f}")

    print()
    print("=== Slope per variant ===")
    by = {}
    for r in results:
        m = re.match(r"O_(\w+?)_Y\d+\.docx", r["file"])
        if not m: continue
        by.setdefault(m.group(1), []).append(r)
    for v in sorted(by):
        rs = sorted(by[v], key=lambda x: x["tblpY_pt"])
        if len(rs) >= 2:
            dy = rs[-1]["table_top_pt"] - rs[0]["table_top_pt"]
            dx = rs[-1]["tblpY_pt"]   - rs[0]["tblpY_pt"]
            slope = dy / dx if dx else None
            tts = [f"{r['table_top_pt']:.2f}" for r in rs]
            print(f"  {v:10s} table_top {tts}  slope={slope:.3f}")


if __name__ == "__main__":
    main()
