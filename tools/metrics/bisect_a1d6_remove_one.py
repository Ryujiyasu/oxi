"""S140 bisection step 2: start from full a1d6, remove ONE table at a time.
If removing table X causes bug to disappear (1 line instead of 2), X is part of the trigger.
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
    out_prefix = os.path.join(TMP, f'r1_{label}')
    out_layout = os.path.join(TMP, f'r1_{label}.json')
    r = subprocess.run([RENDERER, docx_path, out_prefix, f'--dump-layout={out_layout}'],
                       capture_output=True, text=True)
    if r.returncode != 0:
        return None, None
    layout = json.load(open(out_layout, encoding='utf-8'))
    target_cell_para_idx = None
    target_cell_row = None
    target_cell_col = None
    target_page = None
    for pi, page in enumerate(layout['pages']):
        for el in page['elements']:
            if el.get('type') != 'text': continue
            t = el.get('text', '')
            if '役員のうち' in t or 'がある者' in t:
                target_cell_para_idx = el.get('cell_para_idx')
                target_cell_row = el.get('cell_row_idx')
                target_cell_col = el.get('cell_col_idx')
                target_page = pi + 1
                break
        if target_cell_para_idx is not None:
            break
    if target_cell_para_idx is None:
        return None, None
    target_ys = set()
    for page in layout['pages']:
        for el in page['elements']:
            if el.get('type') != 'text': continue
            if (el.get('cell_para_idx') == target_cell_para_idx
                    and el.get('cell_row_idx') == target_cell_row
                    and el.get('cell_col_idx') == target_cell_col):
                target_ys.add(round(el.get('y', 0), 1))
    return len(target_ys), target_page


def main():
    parts = load(SRC)
    xml = parts['word/document.xml'].decode('utf-8')
    tbl_re = re.compile(r'<w:tbl>.*?</w:tbl>', flags=re.DOTALL)
    matches = list(tbl_re.finditer(xml))
    print(f'Found {len(matches)} tables')

    # First, measure original
    n_orig, p_orig = measure(SRC)
    print(f'Original a1d6: 法人等 n_lines={n_orig} on page {p_orig}')

    # Remove each table and remeasure
    print(f'\n=== Remove one table at a time ===')
    print(f'{"variant":<30} | n_lines | page')
    print('-' * 55)
    for k in range(len(matches)):
        new_xml = xml[:matches[k].start()] + xml[matches[k].end():]
        parts_new = dict(parts)
        parts_new['word/document.xml'] = new_xml.encode('utf-8')
        out = os.path.join(OUT_DIR, f'TM_remove_{k}.docx')
        save(parts_new, out)
        n, p = measure(out)
        print(f'{f"remove table {k}":<30} | {n} | {p}')


if __name__ == '__main__':
    main()
