"""Extract d77a tbl5/tbl8/tbl10 XML for structural diff.

Find the tables in document.xml via text search (looking for the starting
content of each), dump cell XML structure for comparison.

Target markers:
- tbl5 first para: "４）　本利用ルールが適用されないコンテンツについて"
- tbl8 first para: "７）　その他" (per cell_paras data: '\\u30b7\\uff09' = '７）')
- tbl10 first para: look up via measurement
"""
import re
import xml.etree.ElementTree as ET
from pathlib import Path


XML = Path(r"C:\tmp\d77a_xml\word\document.xml")
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\d77a_tables_xml_diff.txt")


def main():
    xml = XML.read_text(encoding="utf-8")

    # Find tbl elements by their position in document order.
    # Use a simple regex to enumerate <w:tbl>...</w:tbl> blocks.
    # Match non-greedy to handle nested tables (though d77a's targets are flat).
    # Using regex to split into table blocks.
    # For robustness, look for <w:tbl> at top level and find matching </w:tbl>.

    tables = []
    pos = 0
    while True:
        start = xml.find("<w:tbl>", pos)
        if start == -1:
            break
        # Find matching </w:tbl> by counting depth (nested tables exist)
        depth = 1
        cursor = start + len("<w:tbl>")
        while depth > 0 and cursor < len(xml):
            nx_open = xml.find("<w:tbl>", cursor)
            nx_close = xml.find("</w:tbl>", cursor)
            if nx_close == -1:
                break
            if nx_open != -1 and nx_open < nx_close:
                depth += 1
                cursor = nx_open + len("<w:tbl>")
            else:
                depth -= 1
                cursor = nx_close + len("</w:tbl>")
        end = cursor
        tables.append((start, end, xml[start:end]))
        pos = end

    print(f"Found {len(tables)} tables")

    # Identify by first cell content
    for i, (s, e, body) in enumerate(tables, start=1):
        # Extract first paragraph text (quick preview)
        m = re.search(r"<w:t[^>]*>([^<]{0,60})</w:t>", body)
        preview = m.group(1) if m else "(no text)"
        print(f"Table #{i}: {len(body)} bytes, preview={preview!r}")

    # Targets: tables 5, 8, 10 (1-indexed)
    out_lines = []
    for idx in [5, 8, 10]:
        if idx > len(tables):
            continue
        body = tables[idx - 1][2]
        out_lines.append(f"=== TABLE {idx} ({len(body)} bytes) ===")
        # Pretty-print with indentation for readability
        # Simple approach: split on > and add newlines
        formatted = body.replace("><", ">\n<")
        out_lines.append(formatted)
        out_lines.append("")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("\n".join(out_lines), encoding="utf-8")
    print(f"Wrote {OUT} ({OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
