"""Measure FE_* intervening-empty-paragraph repros."""
import json, re, zipfile, sys
from pathlib import Path
import win32com.client as w32

REPRO_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\fe_repro")
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\fe_intervening_measurements.json")


def main():
    docs = sorted(REPRO_DIR.glob("FE_*.docx"))
    print(f"Measuring {len(docs)} docs...", file=sys.stderr)
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for d in docs:
            with zipfile.ZipFile(d) as z:
                xml = z.read("word/document.xml").decode("utf-8")
            tblpY_list = [int(m.group(1)) for m in re.finditer(r'w:tblpY="(-?\d+)"', xml)]
            doc = word.Documents.Open(str(d.resolve()), ReadOnly=True)
            try:
                # Last table is the TARGET floating table.
                tbl = doc.Tables(doc.Tables.Count)
                tt = tbl.Range.Information(6)
                tp = tbl.Range.Information(3)
                pre_range = doc.Range(0, tbl.Range.Start)
                last_pre = pre_range.Paragraphs(pre_range.Paragraphs.Count)
                ay = last_pre.Range.Information(6)
                anchor_text = (last_pre.Range.Text or "")[:25].replace("\r", "\\r").replace("\x07", "\\x07")
                # Also get y of FIRST body paragraph (the "body para A" anchor).
                first_para = doc.Paragraphs(1)
                fy = first_para.Range.Information(6)
                first_text = (first_para.Range.Text or "")[:25].replace("\r","\\r").replace("\x07","\\x07")
                target_tblpY = tblpY_list[-1] if tblpY_list else 0
                results.append({
                    "file": d.name,
                    "n_tables_word": doc.Tables.Count,
                    "target_tblpY_tw": target_tblpY,
                    "target_tblpY_pt": target_tblpY / 20.0,
                    "first_body_para_y": fy,
                    "first_body_para_text": first_text,
                    "anchor_para_top_pt": ay,
                    "anchor_para_text": anchor_text,
                    "table_top_pt": tt,
                    "table_page": tp,
                    "delta_anchor_to_table": tt - ay if (tt and ay) else None,
                    "y0_intercept_estimate": (tt - ay - target_tblpY/20.0) if (tt and ay) else None,
                })
            finally:
                doc.Close(SaveChanges=0)
            print(f"  done {d.name}", file=sys.stderr)
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    # Summary
    print()
    print(f"{'file':40s} {'#tbl':>4} {'tY_pt':>6} {'1stY':>6} {'aY':>6} {'tY':>6} {'Δ(t-a)':>8} {'Y0int':>7} {'anchor':<20}")
    print("-" * 115)
    for r in results:
        print(f"  {r['file'][:38]:38s}"
              f" {r['n_tables_word']:>4}"
              f" {r['target_tblpY_pt']:>6.2f}"
              f" {r['first_body_para_y']:>6.2f}"
              f" {r['anchor_para_top_pt']:>6.2f}"
              f" {r['table_top_pt']:>6.2f}"
              f" {r['delta_anchor_to_table']:>+8.2f}"
              f" {r['y0_intercept_estimate']:>+7.2f}"
              f" {r['anchor_para_text']!r}")

    # Group by precedent kind
    print()
    print("=== Y0 intercept vs intervening empty paragraphs ===")
    by_pre = {}
    for r in results:
        m = re.match(r"FE_(\w+?)_(\d+)e_Y(\d+)\.docx", r["file"])
        if not m: continue
        pkn, ec, ytw = m.group(1), int(m.group(2)), int(m.group(3))
        by_pre.setdefault((pkn, ytw), []).append((ec, r))
    for (pkn, ytw) in sorted(by_pre):
        rows = sorted(by_pre[(pkn, ytw)])
        print(f"  precedent={pkn:12s} tblpY={ytw:5d}tw  → {[(ec, round(r['y0_intercept_estimate'],2)) for ec,r in rows]}")


if __name__ == "__main__":
    main()
