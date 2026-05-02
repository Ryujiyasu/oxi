"""Measure M_* TP3-mutate repros via Word COM."""
import json, re, zipfile, sys
from pathlib import Path
import win32com.client as w32

REPRO_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\tp3_mutate_repro")
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\tp3_mutate_measurements.json")


def parse_tblpY_tw(p):
    with zipfile.ZipFile(p) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    m = re.search(r'w:tblpY="(-?\d+)"', xml)
    return int(m.group(1)) if m else None


def main():
    docs = sorted(REPRO_DIR.glob("M_*.docx"))
    print(f"Measuring {len(docs)} docs...", file=sys.stderr)
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for d in docs:
            tw = parse_tblpY_tw(d)
            try:
                doc = word.Documents.Open(str(d.resolve()), ReadOnly=True)
            except Exception as e:
                results.append({"file": d.name, "error": str(e)})
                continue
            try:
                tbl = doc.Tables(1)
                tt = tbl.Range.Information(6)
                tp = tbl.Range.Information(3)
                pre_range = doc.Range(0, tbl.Range.Start)
                last_pre = pre_range.Paragraphs(pre_range.Paragraphs.Count)
                ay = last_pre.Range.Information(6)
                results.append({
                    "file": d.name,
                    "tblpY_tw": tw,
                    "tblpY_pt": tw / 20.0 if tw is not None else None,
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
        if "error" in r:
            print(f"  {r['file']:28s}  ERROR: {r['error']}")
            continue
        print(f"  {r['file']:28s}"
              f" {r['tblpY_pt']:>9.2f}"
              f" {r['anchor_top_pt']:>7.2f}"
              f" {r['table_top_pt']:>7.2f}"
              f" {r['delta_pt']:>+7.2f}")

    # Slope per mutation
    print()
    print("=== Slope per mutation ===")
    by_mut = {}
    for r in results:
        if "error" in r:
            continue
        m = re.match(r"M_(\w+?)_Y\d+\.docx", r["file"])
        if not m:
            continue
        by_mut.setdefault(m.group(1), []).append(r)
    for mut in sorted(by_mut):
        rs = sorted(by_mut[mut], key=lambda x: x["tblpY_pt"])
        if len(rs) >= 2:
            dy = rs[-1]["table_top_pt"] - rs[0]["table_top_pt"]
            dx = rs[-1]["tblpY_pt"]   - rs[0]["tblpY_pt"]
            slope = dy / dx if dx else None
            tts = [f"{r['table_top_pt']:.2f}" for r in rs]
            ys  = [f"{r['tblpY_pt']:.2f}" for r in rs]
            print(f"  {mut:14s} tblpY {ys}  table_top {tts}  slope={slope:.3f}")


if __name__ == "__main__":
    main()
