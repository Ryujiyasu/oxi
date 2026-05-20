"""S140 v2: Top-down minimization that specifically counts wrap lines
for the 法人等 paragraph by tracking text characters that originate
from that paragraph's text content.
"""
from __future__ import annotations
import os, sys, re, zipfile, subprocess, json
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
SRC_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'a1d6_top_down_min')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TMP = r'C:\tmp'

# Text fragments unique to the 法人等 paragraph (avoid matching other ○ items)
SIGNATURE_CHARS = list('法人等であって、その役員のうちに上記のいずれかに該当する者がある者')


def measure_oxi_para_lines(docx_path):
    label = os.path.splitext(os.path.basename(docx_path))[0]
    out_prefix = os.path.join(TMP, f'{label}')
    out_layout = os.path.join(TMP, f'{label}_layout.json')
    cmd = [RENDERER, docx_path, out_prefix, f'--dump-layout={out_layout}']
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        return {'error': r.stderr[-500:]}
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    # For each page, find a text element whose text is one of our signature chars
    # Track y values of all signature characters
    target_ys = set()
    target_para_idx = None
    target_cell_para_idx = None
    target_cell_row = None
    target_cell_col = None
    matched_text = None
    matched_page = None
    # First pass: find element whose text uniquely identifies 法人等 paragraph
    for pi, page in enumerate(layout.get('pages', [])):
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            t = el.get('text', '')
            if '役員のうち' in t or 'がある者' in t:
                target_para_idx = el.get('para_idx')
                target_cell_para_idx = el.get('cell_para_idx')
                target_cell_row = el.get('cell_row_idx')
                target_cell_col = el.get('cell_col_idx')
                matched_text = t
                matched_page = pi + 1
                break
        if target_para_idx is not None:
            break
    # Second pass: collect y values for elements matching the same cell+cell_para_idx
    n_target_elements = 0
    if target_para_idx is not None:
        for page in layout.get('pages', []):
            for el in page.get('elements', []):
                if el.get('type') != 'text':
                    continue
                if (el.get('cell_para_idx') == target_cell_para_idx
                        and el.get('cell_row_idx') == target_cell_row
                        and el.get('cell_col_idx') == target_cell_col):
                    target_ys.add(round(el.get('y', 0), 1))
                    n_target_elements += 1
    return {
        'target_para_idx': target_para_idx,
        'target_cell_para_idx': target_cell_para_idx,
        'target_cell_row': target_cell_row,
        'target_cell_col': target_cell_col,
        'matched_text': matched_text,
        'matched_page': matched_page,
        'n_target_elements': n_target_elements,
        'n_lines': len(target_ys),
        'line_ys': sorted(target_ys),
    }


def measure_word(docx_path):
    """Count Word's actual wrap lines for 法人等 paragraph."""
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(docx_path, ReadOnly=True)
    try:
        # Find the 法人等 paragraph
        cell_para = None
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            txt = (p.Range.Text or '').strip()
            if '法人等' in txt and 'であって' in txt:
                cell_para = p
                break
        if cell_para is None:
            return {'error': 'no 法人等 para'}
        rng = cell_para.Range
        line_ys = set()
        n_chars = rng.Characters.Count
        for j in range(1, n_chars + 1):
            ch = rng.Characters(j)
            cr = d.Range(ch.Start, ch.Start)
            yy = round(cr.Information(6), 1)
            line_ys.add(yy)
        return {'n_lines_word': len(line_ys), 'line_ys_word': sorted(line_ys), 'n_chars': n_chars}
    finally:
        d.Close(False)
        word.Quit()


def main():
    variants = ['TM_a_table_only', 'TM_b_row_only', 'TM_c_cell_only', 'TM_d_para_only']
    print(f'{"Variant":<20} | {"Oxi lines":>9} | {"Word lines":>10} | match')
    print('-' * 60)
    for v in variants:
        docx = os.path.join(SRC_DIR, f'{v}.docx')
        if not os.path.exists(docx):
            continue
        o = measure_oxi_para_lines(docx)
        try:
            w = measure_word(docx)
        except Exception as e:
            w = {'error': str(e)[:200]}
        n_o = o.get('n_lines', '?')
        n_w = w.get('n_lines_word', '?')
        match = '✓' if n_o == n_w else '✗'
        print(f'{v:<20} | {n_o:>9} | {n_w:>10} | {match}')
        print(f'    Oxi: matched_text={o.get("matched_text")!r:>10} para_idx={o.get("target_para_idx")} cell=({o.get("target_cell_row")},{o.get("target_cell_col")}) page={o.get("matched_page")} n_target_elements={o.get("n_target_elements")} ys={o.get("line_ys")}')
        print(f'    Word: ys={w.get("line_ys_word")}')


if __name__ == '__main__':
    main()
