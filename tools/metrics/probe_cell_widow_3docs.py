"""Validate Word cell-widow rule on 3 real docs.

For each doc, find cross-page paragraphs that split inside a table cell.
Record the paragraph's widowControl setting (via XML inspection) and whether
Word pushed its first line to the next page (via COM Information).
"""
import win32com.client
import os, sys, json, time

DOCS = [
    r"tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx",
    r"tools/golden-test/documents/docx/1636d28e2c46_tokumei_08_04.docx",
    r"tools/golden-test/documents/docx/d4d126dfe1d9_tokumei_08_01-3.docx",
]

OUT = r"pipeline_data/cell_widow_3docs_probe.json"


def probe(path: str) -> dict:
    fullpath = os.path.abspath(path)
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    result = {"doc": os.path.basename(path), "cross_page_paras": []}
    try:
        doc = word.Documents.Open(fullpath, ReadOnly=True)
        total = doc.Paragraphs.Count
        print(f"[{os.path.basename(path)}] {total} paragraphs", flush=True)
        t0 = time.time()
        for i in range(1, total + 1):
            p = doc.Paragraphs(i).Range
            try:
                pn1 = p.Information(3); y1 = p.Information(6)
                e = doc.Range(max(p.Start, p.End - 1), max(p.Start, p.End - 1))
                pn2 = e.Information(3); y2 = e.Information(6)
            except Exception:
                continue
            # Only cross-page cases
            if pn1 != pn2:
                text = p.Text[:30].replace("\r", " ").replace("\n", " ").replace("\x07", "|")
                widow = None
                in_cell = False
                try:
                    widow = doc.Paragraphs(i).Format.WidowControl  # -1 True, 0 False
                    # Check cell context
                    try:
                        doc.Paragraphs(i).Range.Tables.Count  # raises if not in table
                        in_cell = doc.Paragraphs(i).Range.Information(12)  # wdWithInTable?
                    except Exception:
                        in_cell = False
                except Exception:
                    pass
                result["cross_page_paras"].append({
                    "idx": i, "pn1": pn1, "y1": round(y1,2), "pn2": pn2, "y2": round(y2,2),
                    "text": text, "widow": widow, "in_cell": bool(in_cell),
                })
            if time.time() - t0 > 60:
                print(f"  TIMEOUT at para {i}/{total}", flush=True)
                break
        doc.Close(SaveChanges=False)
    finally:
        word.Quit()
    return result


def main():
    all_data = []
    for p in DOCS:
        d = probe(p)
        print(f"  cross-page paras: {len(d['cross_page_paras'])}")
        for e in d["cross_page_paras"]:
            print(f"    idx={e['idx']:3} p{e['pn1']}→p{e['pn2']} widow={e['widow']} cell={e['in_cell']} {e['text']!r}")
        all_data.append(d)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    print(f"\nSaved: {OUT}")


if __name__ == "__main__":
    main()
