"""b35123: derive paragraph → table/row/cell mapping from OOXML structure.

No Word COM required — uses pure XML parsing to walk the document tree
and tag each paragraph with its (table_idx, row_idx, cell_idx) or None
for body-level paragraphs.

Then cross-reference with Oxi's layout JSON to find which rows have
mis-rendered Y positions.
"""
import os
import sys
import re
import zipfile
import json
from xml.etree import ElementTree as ET

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = "tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx"
OXI_LAYOUT = "pipeline_data/_oxi_b35_layout.txt"
OUT = "pipeline_data/b35123_para_row_map_2026-05-03.json"

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def walk_ooxml():
    """Walk document.xml and tag each w:p with its (table, row, cell) location."""
    with zipfile.ZipFile(DOC) as z:
        xml_bytes = z.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    body = root.find("w:body", NS)
    if body is None:
        print("No body found")
        return []

    paras = []  # (idx, table_idx, row_idx, cell_idx, text_preview)
    para_idx = 0
    table_idx = -1

    def text_of_p(p):
        texts = []
        for t in p.iter(f"{{{NS['w']}}}t"):
            if t.text:
                texts.append(t.text)
        return "".join(texts)

    for child in body:
        tag = child.tag.split("}")[-1]
        if tag == "p":
            para_idx += 1
            paras.append({
                "para_idx": para_idx,
                "location": "body",
                "table_idx": None, "row_idx": None, "cell_idx": None,
                "para_in_cell_idx": None,
                "text_preview": text_of_p(child)[:40],
            })
        elif tag == "tbl":
            table_idx += 1
            row_idx = -1
            for tr in child.findall("w:tr", NS):
                row_idx += 1
                cell_idx = -1
                for tc in tr.findall("w:tc", NS):
                    cell_idx += 1
                    para_in_cell = -1
                    for p in tc.findall("w:p", NS):
                        para_idx += 1
                        para_in_cell += 1
                        paras.append({
                            "para_idx": para_idx,
                            "location": "table_cell",
                            "table_idx": table_idx,
                            "row_idx": row_idx,
                            "cell_idx": cell_idx,
                            "para_in_cell_idx": para_in_cell,
                            "text_preview": text_of_p(p)[:40],
                        })
    return paras


def parse_oxi_layout():
    """Group Oxi layout TEXT entries by (y bucket, para_idx if available).
    Return list of (y, x_start, x_end, text)."""
    if not os.path.exists(OXI_LAYOUT):
        print(f"Oxi layout missing: {OXI_LAYOUT}")
        return []
    lines = []
    cur_y = None
    cur_x = None
    cur_text = ""
    with open(OXI_LAYOUT, encoding="utf-8") as f:
        flines = f.readlines()
    i = 0
    while i < len(flines):
        line = flines[i].rstrip()
        parts = line.split("\t")
        if parts[0] == "TEXT":
            try:
                x = float(parts[1]); y = float(parts[2])
            except (ValueError, IndexError):
                i += 1; continue
            txt = ""
            if i+1 < len(flines) and flines[i+1].startswith("T\t"):
                txt = flines[i+1].rstrip()[2:]
            # Round y to 0.5pt bucket for line grouping
            y_bucket = round(y * 2) / 2
            lines.append({"y_bucket": y_bucket, "x": x, "text": txt})
            i += 2
        else:
            i += 1
    return lines


def main():
    paras = walk_ooxml()
    print(f"Total paragraphs: {len(paras)}", flush=True)

    # Stats by location
    body_count = sum(1 for p in paras if p["location"] == "body")
    cell_count = sum(1 for p in paras if p["location"] == "table_cell")
    print(f"  body: {body_count}, table_cell: {cell_count}", flush=True)

    # Show table 0 + table 1 row breakdown
    tables = {}
    for p in paras:
        if p["table_idx"] is not None:
            tables.setdefault(p["table_idx"], {}).setdefault(p["row_idx"], []).append(p)

    for ti in sorted(tables.keys())[:2]:
        rows = tables[ti]
        print(f"\nTable[{ti}]: {len(rows)} rows", flush=True)
        for ri in sorted(rows.keys()):
            ps = rows[ri]
            cells = {}
            for p in ps:
                cells.setdefault(p["cell_idx"], []).append(p)
            cell_summary = []
            for ci in sorted(cells.keys()):
                cps = cells[ci]
                first_text = cps[0]["text_preview"] if cps else ""
                cell_summary.append(f"c{ci}({len(cps)}p):{first_text[:15]!r}")
            row_paras = [p["para_idx"] for p in ps]
            print(f"  row[{ri}] paras {row_paras[0]}..{row_paras[-1]} ({len(ps)} total): {' | '.join(cell_summary)}",
                  flush=True)

    # Cross-reference with Oxi layout
    oxi_lines = parse_oxi_layout()
    print(f"\nOxi layout: {len(oxi_lines)} text entries (after y-bucket grouping)",
          flush=True)

    # Save
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"paragraphs": paras, "oxi_lines_count": len(oxi_lines)}, f,
                   ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT}", flush=True)


if __name__ == "__main__":
    main()
