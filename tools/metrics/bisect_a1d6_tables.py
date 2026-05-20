"""S140 bisection: keep the 法人等 table + various subsets of other tables to
find which OTHER content triggers the cell-wrap bug.
"""
from __future__ import annotations
import os, sys, re, zipfile, subprocess, json
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
SRC = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', 'a1d6e4efa2e7_tokumei_08_01-4.docx')
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'a1d6_top_down_min')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TMP = r'C:\tmp'


def load(path):
    parts = {}
    with zipfile.ZipFile(path) as zf:
        for n in zf.namelist():
            parts[n] = zf.read(n)
    return parts


def save(parts, dst):
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zf:
        for n, b in parts.items():
            zf.writestr(n, b)


def measure(docx_path):
    label = os.path.splitext(os.path.basename(docx_path))[0]
    out_prefix = os.path.join(TMP, f'bsc_{label}')
    out_layout = os.path.join(TMP, f'bsc_{label}.json')
    r = subprocess.run([RENDERER, docx_path, out_prefix, f'--dump-layout={out_layout}'],
                       capture_output=True, text=True)
    if r.returncode != 0:
        return None
    layout = json.load(open(out_layout, encoding='utf-8'))
    # Find 法人等 paragraph, count its lines
    target_cell_para_idx = None
    target_cell_row = None
    target_cell_col = None
    for page in layout['pages']:
        for el in page['elements']:
            if el.get('type') != 'text': continue
            t = el.get('text', '')
            if '役員のうち' in t or 'がある者' in t:
                target_cell_para_idx = el.get('cell_para_idx')
                target_cell_row = el.get('cell_row_idx')
                target_cell_col = el.get('cell_col_idx')
                break
        if target_cell_para_idx is not None:
            break
    if target_cell_para_idx is None:
        return None
    target_ys = set()
    for page in layout['pages']:
        for el in page['elements']:
            if el.get('type') != 'text': continue
            if (el.get('cell_para_idx') == target_cell_para_idx
                    and el.get('cell_row_idx') == target_cell_row
                    and el.get('cell_col_idx') == target_cell_col):
                target_ys.add(round(el.get('y', 0), 1))
    return len(target_ys)


def main():
    parts = load(SRC)
    xml = parts['word/document.xml'].decode('utf-8')

    # Find all <w:tbl>...</w:tbl> blocks
    tbl_re = re.compile(r'<w:tbl>.*?</w:tbl>', flags=re.DOTALL)
    matches = list(tbl_re.finditer(xml))
    print(f'Found {len(matches)} tables in document')

    # Identify the table containing 法人等
    target_tbl_idx = None
    for i, m in enumerate(matches):
        if '法人等であって' in m.group(0):
            target_tbl_idx = i
            break
    print(f'Target table (contains 法人等): index {target_tbl_idx}')

    # Try each pair: target + 1 other table
    body_start = xml.find('<w:body>') + len('<w:body>')
    body_end = xml.find('</w:body>')
    sect_pr = xml[xml.find('<w:sectPr>', body_start):body_end]

    print(f'\n=== Pairwise test: target table + 1 other ===')
    print(f'{"variant":<30} | n_lines')
    print('-' * 50)

    target_table_xml = matches[target_tbl_idx].group(0)

    for j in range(len(matches)):
        if j == target_tbl_idx:
            continue
        other_xml = matches[j].group(0)
        # Build body with [other if j < target else target_first, then the other]
        if j < target_tbl_idx:
            body = other_xml + '\n' + target_table_xml + '\n' + sect_pr
        else:
            body = target_table_xml + '\n' + other_xml + '\n' + sect_pr
        new_xml = xml[:body_start] + body + xml[body_end:]
        parts_new = dict(parts)
        parts_new['word/document.xml'] = new_xml.encode('utf-8')
        out = os.path.join(OUT_DIR, f'TM_pair_{target_tbl_idx}_{j}.docx')
        save(parts_new, out)
        n_lines = measure(out)
        print(f'{f"target + table {j}":<30} | {n_lines}')

    # Also test target alone (sanity)
    body_alone = target_table_xml + '\n' + sect_pr
    new_xml = xml[:body_start] + body_alone + xml[body_end:]
    parts_new = dict(parts)
    parts_new['word/document.xml'] = new_xml.encode('utf-8')
    out = os.path.join(OUT_DIR, f'TM_target_alone.docx')
    save(parts_new, out)
    print(f'{"target alone":<30} | {measure(out)}')


if __name__ == '__main__':
    main()
