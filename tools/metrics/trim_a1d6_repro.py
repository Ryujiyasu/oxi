"""S140: Top-down minimization from a1d6 to isolate cell-wrap bug context.

Approach: take a1d6 docx, progressively strip content from its document.xml,
re-pack as new docx, render with Oxi + compare with Word.

Variants:
  TM_a: keep only table containing 法人等 (drop other tables)
  TM_b: TM_a + keep only the row containing 法人等
  TM_c: TM_b + keep only the cell with 法人等
  TM_d: TM_c + keep only the 法人等 paragraph (drop other ○ siblings)
  TM_e: TM_d + remove all other body content

For each variant, check if Oxi still over-wraps 法人等 to 2 lines.
Smallest variant that still over-wraps = the bug-triggering context.
"""
from __future__ import annotations
import os, sys, re, zipfile, shutil, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
SRC = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', 'a1d6e4efa2e7_tokumei_08_01-4.docx')
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'a1d6_top_down_min')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TMP = r'C:\tmp'


def load_docx(path):
    parts = {}
    with zipfile.ZipFile(path) as zf:
        for n in zf.namelist():
            parts[n] = zf.read(n)
    return parts


def save_docx(parts, dst):
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zf:
        for n, b in parts.items():
            zf.writestr(n, b)


def find_table_containing(xml, marker):
    """Find <w:tbl>...</w:tbl> containing the marker text. Returns (start, end)."""
    idx = xml.find(marker)
    if idx < 0:
        return None
    # Find enclosing <w:tbl>
    t_start = xml.rfind('<w:tbl>', 0, idx)
    if t_start < 0:
        t_start = xml.rfind('<w:tbl ', 0, idx)
    t_end = xml.find('</w:tbl>', idx) + len('</w:tbl>')
    return (t_start, t_end)


def find_row_containing(xml, marker, tbl_range):
    s, e = tbl_range
    idx = xml.find(marker, s, e)
    r_start = xml.rfind('<w:tr ', 0, idx)
    if r_start < 0 or r_start < s:
        r_start = xml.rfind('<w:tr>', 0, idx)
    r_end = xml.find('</w:tr>', idx) + len('</w:tr>')
    return (r_start, r_end)


def find_cell_containing(xml, marker, row_range):
    s, e = row_range
    idx = xml.find(marker, s, e)
    c_start = xml.rfind('<w:tc>', 0, idx)
    c_end = xml.find('</w:tc>', idx) + len('</w:tc>')
    return (c_start, c_end)


def find_para_containing(xml, marker, cell_range):
    s, e = cell_range
    idx = xml.find(marker, s, e)
    p_start = xml.rfind('<w:p ', 0, idx)
    if p_start < 0:
        p_start = xml.rfind('<w:p>', 0, idx)
    p_end = xml.find('</w:p>', idx) + len('</w:p>')
    return (p_start, p_end)


def measure_oxi(docx_path):
    """Render and find the y count for the 法人等 cell's text elements."""
    import json
    label = os.path.splitext(os.path.basename(docx_path))[0]
    out_prefix = os.path.join(TMP, f'{label}')
    out_layout = os.path.join(TMP, f'{label}_layout.json')
    cmd = [RENDERER, docx_path, out_prefix, f'--dump-layout={out_layout}']
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        return {'error': r.stderr[-500:]}
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    # Count unique y values for text elements containing 法人 or in_table cells with relevant text
    target_ys = set()
    for page in layout.get('pages', []):
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            if el.get('cell_row_idx') is None:
                continue
            # Match all in-cell text on same page in the target cell
            # Use text-prefix matching for the cell we care about
            t = el.get('text', '')
            # Identify text element that's part of the 法人等 paragraph
            # Look for any character of the paragraph
            target_ys.add(round(el.get('y', 0), 1))
    return {'n_lines_oxi_all_cell_text': len(target_ys), 'lines': sorted(target_ys)}


def measure_oxi_para(docx_path, prev_text_y_set):
    """Same as measure_oxi but compare against a prior 'before' set to find delta."""
    return measure_oxi(docx_path)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    parts = load_docx(SRC)
    doc_xml_bytes = parts['word/document.xml']
    xml = doc_xml_bytes.decode('utf-8')
    marker = '法人等であって'  # the trigger paragraph

    # First, find structures
    tbl = find_table_containing(xml, marker)
    if not tbl:
        print(f'marker not found')
        return
    print(f'tbl range: {tbl[0]}..{tbl[1]} (length {tbl[1] - tbl[0]})')
    row = find_row_containing(xml, marker, tbl)
    print(f'row range: {row[0]}..{row[1]} (length {row[1] - row[0]})')
    cell = find_cell_containing(xml, marker, row)
    print(f'cell range: {cell[0]}..{cell[1]} (length {cell[1] - cell[0]})')
    para = find_para_containing(xml, marker, cell)
    print(f'para range: {para[0]}..{para[1]} (length {para[1] - para[0]})')

    # Body start/end
    body_start = xml.find('<w:body>') + len('<w:body>')
    body_end = xml.find('</w:body>')
    sect_pr = xml[xml.find('<w:sectPr>', body_start):body_end]

    # TM_a: keep only the table (and sectPr)
    body_a = xml[tbl[0]:tbl[1]] + '\n' + sect_pr
    new_xml = xml[:body_start] + body_a + xml[body_end:]
    parts_a = dict(parts)
    parts_a['word/document.xml'] = new_xml.encode('utf-8')
    out_a = os.path.join(OUT_DIR, 'TM_a_table_only.docx')
    save_docx(parts_a, out_a)
    print(f'\nTM_a: only table containing 法人等')
    r = measure_oxi(out_a)
    print(f'  Oxi: {r}')

    # TM_b: keep only the row (need a minimal table wrapper)
    # Find table opening (tblPr + tblGrid) and use that
    tbl_xml = xml[tbl[0]:tbl[1]]
    # Extract <w:tblPr>...</w:tblPr> and <w:tblGrid>...</w:tblGrid>
    tbl_pr_m = re.search(r'<w:tblPr>.*?</w:tblPr>', tbl_xml, flags=re.DOTALL)
    tbl_grid_m = re.search(r'<w:tblGrid>.*?</w:tblGrid>', tbl_xml, flags=re.DOTALL)
    tbl_pr = tbl_pr_m.group(0) if tbl_pr_m else ''
    tbl_grid = tbl_grid_m.group(0) if tbl_grid_m else ''
    row_xml = xml[row[0]:row[1]]
    body_b = f'<w:tbl>{tbl_pr}{tbl_grid}{row_xml}</w:tbl>\n' + sect_pr
    new_xml = xml[:body_start] + body_b + xml[body_end:]
    parts_b = dict(parts)
    parts_b['word/document.xml'] = new_xml.encode('utf-8')
    out_b = os.path.join(OUT_DIR, 'TM_b_row_only.docx')
    save_docx(parts_b, out_b)
    print(f'\nTM_b: only row containing 法人等')
    r = measure_oxi(out_b)
    print(f'  Oxi: {r}')

    # TM_c: only the cell with 法人等 — replace row to have just this 1 cell
    # Keep <w:tr> opening (preserve trPr if any)
    tr_open_m = re.match(r'<w:tr[^>]*>', xml[row[0]:row[1]])
    tr_open = tr_open_m.group(0) if tr_open_m else '<w:tr>'
    cell_xml = xml[cell[0]:cell[1]]
    body_c = (f'<w:tbl>{tbl_pr}{tbl_grid}'
              f'{tr_open}{cell_xml}</w:tr>'
              '</w:tbl>\n' + sect_pr)
    new_xml = xml[:body_start] + body_c + xml[body_end:]
    parts_c = dict(parts)
    parts_c['word/document.xml'] = new_xml.encode('utf-8')
    out_c = os.path.join(OUT_DIR, 'TM_c_cell_only.docx')
    save_docx(parts_c, out_c)
    print(f'\nTM_c: only cell with 法人等')
    r = measure_oxi(out_c)
    print(f'  Oxi: {r}')

    # TM_d: only the para inside cell
    para_xml = xml[para[0]:para[1]]
    # Extract <w:tc> opening (with tcPr)
    tc_open_m = re.search(r'<w:tc>.*?</w:tcPr>', cell_xml, flags=re.DOTALL)
    tc_open = tc_open_m.group(0) if tc_open_m else '<w:tc>'
    body_d = (f'<w:tbl>{tbl_pr}{tbl_grid}'
              f'{tr_open}{tc_open}{para_xml}</w:tc></w:tr>'
              '</w:tbl>\n' + sect_pr)
    new_xml = xml[:body_start] + body_d + xml[body_end:]
    parts_d = dict(parts)
    parts_d['word/document.xml'] = new_xml.encode('utf-8')
    out_d = os.path.join(OUT_DIR, 'TM_d_para_only.docx')
    save_docx(parts_d, out_d)
    print(f'\nTM_d: only the 法人等 paragraph in cell')
    r = measure_oxi(out_d)
    print(f'  Oxi: {r}')


if __name__ == '__main__':
    main()
