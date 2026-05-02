"""Cluster B verification: long-doc cumulative Y drift hypothesis.

For e3c545 (541 paras) / 04b88e (386 paras) / 34140b (499 paras),
sample paragraph Y at strategic positions:
- First 5 paragraphs (drift baseline)
- Mid-doc paragraphs (5 around middle)
- Last 5 paragraphs (cumulative drift endpoint)
- Plus paragraphs at each page boundary

Compare Word COM-measured Y vs Oxi cached layout.json (if available).

If Oxi has growing |dy| from p1 → p_max, drift is real.
If |dy| is flat, drift is per-page (page break resets).
"""
import json
import sys
import time
from pathlib import Path
import pythoncom
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCS = [
    {"path": "tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx",
     "id": "e3c545"},
    {"path": "tools/golden-test/documents/docx/04b88e7e0b25_index-19.docx",
     "id": "04b88"},
    {"path": "tools/golden-test/documents/docx/34140b9c5662_index-14.docx",
     "id": "34140b"},
]

OUT = Path("pipeline_data/long_doc_drift_measurement.json")


def measure(word, doc_path, doc_id):
    last = None
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(Path(doc_path).resolve()), ReadOnly=True)
            time.sleep(0.5)
            n_paras = doc.Paragraphs.Count
            # Sample positions: first 5, last 5, every 50 in between
            sample_idx = sorted(set(
                list(range(1, min(6, n_paras + 1)))
                + list(range(50, n_paras + 1, 50))
                + list(range(max(1, n_paras - 4), n_paras + 1))
            ))
            samples = []
            for i in sample_idx:
                try:
                    p = doc.Paragraphs(i)
                    rng = p.Range
                    y = rng.Information(6)  # vertical position
                    page = rng.Information(3)  # page number
                    txt = (rng.Text or "")[:30].replace("\r", "\\r").replace("\x07", "\\x07")
                    samples.append({"i": i, "y": round(y, 2), "page": page, "text": txt})
                except Exception:
                    samples.append({"i": i, "error": "info_failed"})
            doc.Close(SaveChanges=False)
            return {"doc_id": doc_id, "n_paras": n_paras, "samples": samples}
        except Exception as e:
            last = e
            time.sleep(0.8 + attempt * 0.5)
    return {"doc_id": doc_id, "error": str(last)}


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    try:
        for d in DOCS:
            r = measure(word, d["path"], d["id"])
            results.append(r)
            if "error" in r:
                print(f"{d['id']}: ERR {r['error']}")
            else:
                print(f"{d['id']}: n_paras={r['n_paras']}, sampled {len(r['samples'])}")
                # print sample row
                for s in r["samples"][:3] + r["samples"][-3:]:
                    if "error" not in s:
                        print(f"    p{s['i']:>4}  page={s['page']}  y={s['y']:>7.1f}  {s['text']!r}")
    finally:
        try: word.Quit()
        except: pass

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
