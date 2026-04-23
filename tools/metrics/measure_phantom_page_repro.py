"""Measure Word's layout of RPH_1..6 phantom-page repros.

For each docx, dump page_count and every paragraph's (idx, page, y, text).
Look for: does Word ever create a phantom page (a page with only empties)?
"""
import json
from pathlib import Path
import win32com.client as w32


REPRO_DIR = Path(__file__).parent / "phantom_page_repro"
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\phantom_page_measurements.json")


def measure(doc_path: Path):
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    result = {"file": doc_path.name}
    try:
        doc = word.Documents.Open(str(doc_path.resolve()), ReadOnly=True)
        try:
            result["page_count"] = doc.ComputeStatistics(2)
            paras = []
            for i, p in enumerate(doc.Paragraphs, start=1):
                r = p.Range
                try:
                    y = r.Information(6)
                    pg = r.Information(3)
                except Exception:
                    continue
                t = r.Text[:80].replace("\r", "\\r").replace("\x07", "\\x07").replace("\n", "\\n")
                is_empty = (r.Text.strip("\r\n\x07\t ") == "")
                paras.append({
                    "idx": i, "page": pg,
                    "y_pt": round(y, 3),
                    "is_empty": is_empty,
                    "text": t,
                })
            result["paras"] = paras
        finally:
            doc.Close(SaveChanges=0)
    finally:
        word.Quit()
    return result


def main():
    results = {}
    for docx in sorted(REPRO_DIR.glob("*.docx")):
        print(f"Measuring {docx.name}...")
        r = measure(docx)
        results[docx.name] = r
        print(f"  pages={r['page_count']}, paras={len(r['paras'])}")
        for p in r["paras"]:
            mark = "[E]" if p["is_empty"] else "[B]"
            print(f"    pg={p['page']} y={p['y_pt']:6.2f} {mark} {p['text']}")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
