"""Re-measure existing TP1-6 repros to verify the prior slope=0 observation.
Prior data: tools/metrics/tblppr_anchor_repro/TP{1..6}.docx all returned
table_top=71.0 (TP1-3) or 98.0 (TP4-6) regardless of tblpY.

Output: pipeline_data/tp_resweep_measurements.json
"""
import json
import re
import zipfile
from pathlib import Path
import sys

import win32com.client as w32

REPRO_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\tblppr_anchor_repro")
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\tp_resweep_measurements.json")


def parse_tblpY(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    m = re.search(r'<w:tblpPr\b[^/>]*?w:tblpY="(-?\d+)"', xml)
    return int(m.group(1)) if m else None


def main():
    docs = sorted(REPRO_DIR.glob("TP*.docx"))
    print(f"Re-measuring {len(docs)} TP repros...", file=sys.stderr)
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for d in docs:
            tw = parse_tblpY(d)
            doc = word.Documents.Open(str(d.resolve()), ReadOnly=True)
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
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)

    print()
    print(f"{'file':40s} {'tblpY_tw':>9} {'tblpY_pt':>9} {'a_top':>7} {'t_top':>7} {'delta':>7}")
    print("-" * 90)
    for r in results:
        print(f"  {r['file']:38s}"
              f" {r['tblpY_tw']:>9}"
              f" {r['tblpY_pt']:>9.2f}"
              f" {r['anchor_top_pt']:>7.2f}"
              f" {r['table_top_pt']:>7.2f}"
              f" {r['delta_pt']:>+7.2f}")


if __name__ == "__main__":
    main()
