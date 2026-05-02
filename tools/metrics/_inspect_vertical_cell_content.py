# -*- coding: utf-8 -*-
"""Find vertical cells in 459f05 and dump their content."""
import sys, os, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import sys
fname = sys.argv[1] if len(sys.argv) > 1 else '459f05f1e877_kyodokenkyuyoushiki01.docx'
DOCX = rf'C:\Users\ryuji\oxi-4\tools\golden-test\documents\docx\{fname}'
print(f"Inspecting: {fname}\n")

with zipfile.ZipFile(DOCX) as z:
    doc = z.read('word/document.xml').decode('utf-8')

# Find each tcPr with textDirection tbRlV, then extract the cell content
# Strategy: find <w:tc> elements containing <w:textDirection w:val="tbRlV"
print("=== Searching for vertical cells (tbRlV) ===")
tc_pattern = re.compile(r'<w:tc>.*?</w:tc>', re.DOTALL)
vertical_cells = []
for m in tc_pattern.finditer(doc):
    cell_xml = m.group(0)
    if 'tbRlV' in cell_xml:
        # Extract text
        texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', cell_xml)
        text_concat = ''.join(texts)
        # Extract pPr/rPr
        td_match = re.search(r'<w:textDirection[^/>]*?w:val="([^"]+)"', cell_xml)
        # Cell width
        tcw = re.search(r'<w:tcW\s+w:w="(\d+)"\s+w:type="(\w+)"', cell_xml)
        # Cell vAlign
        valign = re.search(r'<w:vAlign[^/>]*?w:val="(\w+)"', cell_xml)
        vertical_cells.append({
            'text': text_concat,
            'textDir': td_match.group(1) if td_match else None,
            'tcW': f"{tcw.group(1)}/{tcw.group(2)}" if tcw else None,
            'vAlign': valign.group(1) if valign else None,
            'pos': m.start(),
            'len': len(cell_xml),
        })

print(f"Found {len(vertical_cells)} vertical cells")
for i, c in enumerate(vertical_cells):
    print(f"\n[{i+1}] pos={c['pos']} len={c['len']}b")
    print(f"  textDir: {c['textDir']}")
    print(f"  tcW: {c['tcW']}, vAlign: {c['vAlign']}")
    print(f"  text: {c['text'][:80]!r}")

# Also find non-vertical cells in same table for comparison
print("\n=== Tables with vertical cells ===")
tbl_pattern = re.compile(r'<w:tbl>.*?</w:tbl>', re.DOTALL)
for tm in tbl_pattern.finditer(doc):
    tbl_xml = tm.group(0)
    if 'tbRlV' in tbl_xml:
        # Count cells, rows
        n_rows = len(re.findall(r'<w:tr[\s>]', tbl_xml))
        n_cells = len(re.findall(r'<w:tc>', tbl_xml))
        print(f"Table at pos {tm.start()}: {n_rows} rows, {n_cells} cells")
        # Show table grid widths
        grid = re.search(r'<w:tblGrid>(.*?)</w:tblGrid>', tbl_xml, re.DOTALL)
        if grid:
            cols = re.findall(r'<w:gridCol\s+w:w="(\d+)"', grid.group(1))
            print(f"  tblGrid: {cols}")
