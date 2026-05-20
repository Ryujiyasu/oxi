"""S152: Measure V500 cell-context empty-para + section header repros."""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'cell_empty_section')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

VARIANTS = [
    'V500a_cell_section_only',
    'V500b_cell_1empty_section',
    'V500c_cell_5empty_section',
    'V500d_cell_5empty_section_valign_center',
    'V500e_cell_5empty_section_trheight200',
    'V500f_cell_5empty_section_trheight3000',
    'V500g_cell_5empty_section_no_sb',
    'V500h_cell_5bare_empty_section',
    'V500i_cell_1empty_ac_section',
]


def measure_word(docx_path: str) -> dict:
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(docx_path, ReadOnly=True)
    paras = []
    try:
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            txt = (p.Range.Text or '').strip()
            rng = p.Range
            cr = d.Range(rng.Start, rng.Start)
            paras.append({
                'i': i,
                'text': txt[:30],
                'y': round(cr.Information(6), 3),
                'in_table': bool(cr.Information(12)),
            })
    finally:
        d.Close(False)
        word.Quit()
    return {'paras': paras}


def measure_oxi(docx_path: str) -> dict:
    label = os.path.splitext(os.path.basename(docx_path))[0]
    out_layout = os.path.join(r'C:\tmp', f'{label}.json')
    r = subprocess.run([RENDERER, docx_path, os.path.join(r'C:\tmp', label), f'--dump-layout={out_layout}'],
                       capture_output=True, text=True)
    if r.returncode != 0:
        return {'error': r.stderr[-300:]}
    layout = json.load(open(out_layout, encoding='utf-8'))
    # Find section header (with 取り扱う or SECTION_NO_SB) and BEFORE/AFTER
    section_y, before_y, after_y = None, None, None
    for page in layout['pages']:
        for el in page['elements']:
            if el.get('type') != 'text': continue
            t = el.get('text', '')
            if section_y is None and ('取り扱う' in t or 'SECTION_NO_SB' in t):
                section_y = el['y']
            if before_y is None and 'BEFORE' in t:
                before_y = el['y']
            if after_y is None and 'AFTER' in t:
                after_y = el['y']
    return {'section_y': section_y, 'before_y': before_y, 'after_y': after_y}


def main():
    print(f'{"Variant":<45} | {"BEFORE":>7} | {"section":>14} | {"AFTER":>7} | Word-Oxi dy')
    print('-' * 105)
    for label in VARIANTS:
        docx = os.path.join(REPRO_DIR, f'{label}.docx')
        if not os.path.exists(docx):
            continue
        try:
            w = measure_word(docx)
        except Exception as e:
            print(f'{label}: Word ERROR: {e}')
            continue
        o = measure_oxi(docx)
        if 'error' in o:
            print(f'{label}: Oxi ERROR: {o["error"]}')
            continue
        # Find Word's section header y position
        w_section = next((p for p in w['paras'] if '取り扱う' in p['text'] or 'SECTION_NO_SB' in p['text']), None)
        w_before = next((p for p in w['paras'] if 'BEFORE' in p['text']), None)
        w_after = next((p for p in w['paras'] if 'AFTER' in p['text']), None)
        if not w_section or not o.get('section_y'):
            print(f'{label}: missing section')
            continue
        dy = o['section_y'] - w_section['y']
        w_b = w_before['y'] if w_before else 0
        w_s = w_section['y']
        w_a = w_after['y'] if w_after else 0
        o_b = o.get('before_y', 0) or 0
        o_s = o.get('section_y', 0) or 0
        o_a = o.get('after_y', 0) or 0
        print(f'{label:<45} | W={w_b:>5}/O={o_b:>5} | W={w_s:>5}/O={o_s:>5} ({w_section["in_table"]}) | W={w_a:>5}/O={o_a:>5} | {dy:+.2f}')


if __name__ == '__main__':
    main()
