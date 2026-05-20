"""S139: Measure V300a-V300l minimal repros — Word vs Oxi line count.

For each variant:
- Word COM: count rendered lines in the cell paragraph (via
  range bookmark navigation OR by measuring para height / line height).
- Oxi: dump --dump-layout and count distinct y values for text elements
  in the cell paragraph (= rendered line count).

Output: pipeline_data/cell_wrap_hanging_results.json
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'cell_wrap_hanging')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TMP = r'C:\tmp'

VARIANTS = [
    'V300a_no_hanging',
    'V300b_a1d6_exact',
    'V300c_wider_cell',
    'V300d_sz18',
    'V300e_spacing_neg1',
    'V300f_jc_both',
    'V300g_tcmar0',
    'V300h_hanging50',
    'V300i_no_marker',
    'V300j_text27',
    'V300k_text40',
    'V300l_no_pstyle',
    'V300m_multi_run',
]


def measure_word(docx_path: str) -> dict:
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(docx_path, ReadOnly=True)
    out = {}
    try:
        # First paragraph in table (the one we built)
        # Iterate paragraphs, find the one with our cell text
        cell_para = None
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            txt = (p.Range.Text or '').strip()
            if '法人等' in txt:
                cell_para = p
                break
        if cell_para is None:
            out['error'] = 'cell para not found'
            return out
        rng = cell_para.Range
        cr = d.Range(rng.Start, rng.Start)
        out['cell_para_y'] = round(cr.Information(6), 3)
        out['cell_para_x'] = round(cr.Information(5), 3)
        # Para height via end Y vs start Y (only valid if same page)
        cre = d.Range(rng.End - 1, rng.End - 1)
        out['cell_para_end_y'] = round(cre.Information(6), 3)
        # Words.Count gives word count; for char count use rng.Characters.Count
        out['n_chars'] = len(txt)
        # Get table-row Height for verification
        if d.Tables.Count >= 1:
            t = d.Tables(1)
            r = t.Rows(1)
            out['row_height'] = round(r.Height, 3)
            out['row_height_rule'] = int(r.HeightRule)
            try:
                out['cell_height'] = round(t.Cell(1, 1).Height, 3)
            except Exception:
                out['cell_height'] = None
        # Count visible lines: scan all characters' y positions
        line_ys = set()
        # rng.Characters can be slow but for ~40 chars OK
        for j in range(1, rng.Characters.Count + 1):
            ch = rng.Characters(j)
            chr_rng = d.Range(ch.Start, ch.Start)
            yy = round(chr_rng.Information(6), 1)
            line_ys.add(yy)
        out['n_lines_word'] = len(line_ys)
        out['line_ys_word'] = sorted(line_ys)
    finally:
        d.Close(False)
        word.Quit()
    return out


def measure_oxi(docx_path: str) -> dict:
    label = os.path.splitext(os.path.basename(docx_path))[0]
    out_prefix = os.path.join(TMP, f'{label}')
    out_layout = os.path.join(TMP, f'{label}_layout.json')
    cmd = [RENDERER, docx_path, out_prefix, f'--dump-layout={out_layout}']
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        return {'error': r.stderr[-500:]}
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    # Find text elements that are in a cell with text containing 法人等
    target_ys = set()
    for page in layout.get('pages', []):
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            if el.get('cell_row_idx') is None:
                continue
            txt = el.get('text', '')
            # Match if any text element in this cell contains 法人等 — but
            # we don't have per-cell index aggregation here. Use simpler approach:
            # match all text in the first row of first table (only one row in our repros).
            target_ys.add(round(el.get('y', 0), 1))
    return {
        'n_lines_oxi': len(target_ys),
        'line_ys_oxi': sorted(target_ys),
    }


def main():
    results = []
    for label in VARIANTS:
        docx = os.path.join(REPRO_DIR, f'{label}.docx')
        if not os.path.exists(docx):
            print(f'SKIP {label}: not found')
            continue
        print(f'=== {label} ===')
        try:
            w = measure_word(docx)
        except Exception as e:
            print(f'  Word ERROR: {e}')
            results.append({'label': label, 'word_error': str(e)})
            continue
        try:
            o = measure_oxi(docx)
        except Exception as e:
            print(f'  Oxi ERROR: {e}')
            results.append({'label': label, 'word': w, 'oxi_error': str(e)})
            continue
        if 'error' in o:
            print(f'  Oxi render error: {o["error"]}')
            results.append({'label': label, 'word': w, 'oxi_error': o['error']})
            continue
        n_w = w.get('n_lines_word')
        n_o = o.get('n_lines_oxi')
        match = '✓' if n_w == n_o else '✗ MISMATCH'
        print(f'  n_chars={w.get("n_chars")} '
              f'Word lines={n_w} Oxi lines={n_o} {match} '
              f'row_h_word={w.get("row_height")} '
              f'cell_y={w.get("cell_para_y")}')
        results.append({
            'label': label,
            'word': w,
            'oxi': o,
            'match': (n_w == n_o),
        })
    out_path = os.path.join(REPO, 'pipeline_data', 'cell_wrap_hanging_results.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'\nWrote {out_path}')


if __name__ == '__main__':
    main()
