"""For each table row in a docx, dump where <w:lastRenderedPageBreak/> markers
appear inside the row's content.

Output per LRPB hit:
  doc_id, table_index, row_index, cell_index, para_index_in_cell, run_index_in_cell_para, paragraph_text_prefix

Used in Day 35 session 58+ to discriminate Word's push-vs-split row decisions:
- Row with no LRPB inside it: not broken at this row
- Row with LRPB at (cell=0, para=0, run=0): row was PUSHED whole to next page
- Row with LRPB at run>0 OR para>0 OR cell>0: row was SPLIT across pages (break point is inside)

Usage: python lrpb_table_positions.py <doc_id> [<doc_id> ...]
"""
import os, sys, zipfile, json
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOCX_DIR = os.path.join(REPO_ROOT, "tools", "golden-test", "documents", "docx")


def find_docx(doc_id):
    for f in os.listdir(DOCX_DIR):
        if f.startswith(doc_id) and f.endswith(".docx"):
            return os.path.join(DOCX_DIR, f)
    return None


def text_prefix(elem, n=40):
    """Concatenate all w:t text within an element."""
    parts = []
    for t in elem.iter(W + "t"):
        if t.text:
            parts.append(t.text)
    s = "".join(parts)
    return s[:n]


def scan_doc(doc_id):
    path = find_docx(doc_id)
    if not path:
        return {"doc_id": doc_id, "error": "docx not found", "rows": []}
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    body = root.find(W + "body")
    out_rows = []
    table_idx = -1
    for elem in body:
        if elem.tag == W + "tbl":
            table_idx += 1
            # is floating?
            tblpr = elem.find(W + "tblPr")
            floating = False
            if tblpr is not None and tblpr.find(W + "tblpPr") is not None:
                floating = True
            row_idx = -1
            for tr in elem.findall(W + "tr"):
                row_idx += 1
                row_info = {
                    "table_index": table_idx,
                    "floating": floating,
                    "row_index": row_idx,
                    "cell_count": 0,
                    "lrpb_hits": [],
                    "row_text_prefix": "",
                }
                cell_idx = -1
                texts_for_row = []
                for tc in tr.findall(W + "tc"):
                    cell_idx += 1
                    para_idx = -1
                    for p in tc.findall(W + "p"):
                        para_idx += 1
                        run_idx = -1
                        for r in p.findall(W + "r"):
                            run_idx += 1
                            if r.find(W + "lastRenderedPageBreak") is not None:
                                row_info["lrpb_hits"].append({
                                    "cell_index": cell_idx,
                                    "para_index_in_cell": para_idx,
                                    "run_index_in_cell_para": run_idx,
                                    "para_text_prefix": text_prefix(p),
                                })
                        # Also collect text for row preview
                        if cell_idx == 0 and para_idx == 0:
                            tp = text_prefix(p, 30)
                            if tp:
                                texts_for_row.append(tp)
                row_info["cell_count"] = cell_idx + 1
                row_info["row_text_prefix"] = " | ".join(texts_for_row)[:60]
                # Only include rows with LRPB hits OR all rows? Keep all for now
                if row_info["lrpb_hits"]:
                    out_rows.append(row_info)
    return {"doc_id": doc_id, "n_rows_with_lrpb": len(out_rows), "rows": out_rows}


def main():
    docs = sys.argv[1:]
    if not docs:
        docs = ["29dc6e8943fe", "31420af1a08f", "6514f214e482", "de6e32b5960b",
                "d4d126dfe1d9", "459f05f1e877"]
    for doc_id in docs:
        print(f"=== {doc_id} ===")
        result = scan_doc(doc_id)
        if "error" in result:
            print(f"  ERROR: {result['error']}")
            continue
        print(f"  rows with LRPB: {result['n_rows_with_lrpb']}")
        for row in result["rows"]:
            float_mark = " [FLOAT]" if row["floating"] else ""
            print(f"  table={row['table_index']}{float_mark} row={row['row_index']} cells={row['cell_count']}")
            print(f"    row_text: {row['row_text_prefix']!r}")
            for hit in row["lrpb_hits"]:
                print(f"    LRPB @ cell={hit['cell_index']} para={hit['para_index_in_cell']} run={hit['run_index_in_cell_para']}")
                print(f"      para_text: {hit['para_text_prefix']!r}")
        print()


if __name__ == "__main__":
    main()
