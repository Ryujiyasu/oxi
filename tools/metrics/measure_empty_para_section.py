"""S150: Measure TR_V400 empty-para-section repros.
Compare Word's section header y position vs Oxi's.
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'empty_para_section')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

VARIANTS = [
    'V400a_empty_default_then_section',
    'V400b_no_empty_para',
    'V400c_two_empty_paras',
    'V400d_empty_exact240',
    'V400e_empty_pstyle_ac',
    'V400f_no_beforeLines',
    'V400g_empty_auto0',
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
    paras = {}
    for page in layout['pages']:
        for el in page['elements']:
            if el.get('type') != 'text': continue
            pi = el.get('para_idx')
            if pi is None: continue
            paras.setdefault(pi, []).append((el['y'], el['x'], el['text']))
    # Build sorted list
    out = []
    for pi in sorted(paras.keys()):
        elems = sorted(paras[pi])
        text = ''.join(t[2] for t in elems)
        out.append({'i': pi, 'text': text[:30], 'y': elems[0][0] if elems else 0})
    return {'paras': out}


def main():
    print(f'{"Variant":<35} | {"BEFORE":>7} | {"section":>7} | {"AFTER":>7} | sec-bef | Word-Oxi sec')
    print('-' * 100)
    for label in VARIANTS:
        docx = os.path.join(REPRO_DIR, f'{label}.docx')
        if not os.path.exists(docx):
            print(f'{label}: NOT FOUND')
            continue
        try:
            w = measure_word(docx)
        except Exception as e:
            print(f'{label}: Word ERROR: {e}')
            continue
        try:
            o = measure_oxi(docx)
        except Exception as e:
            print(f'{label}: Oxi ERROR: {e}')
            continue

        # Find BEFORE, section header, AFTER in each
        w_before = next((p for p in w['paras'] if 'BEFORE' in p['text']), None)
        w_section = next((p for p in w['paras'] if '匿名データの提供' in p['text']), None)
        w_after = next((p for p in w['paras'] if 'AFTER' in p['text']), None)
        o_before = next((p for p in o['paras'] if 'BEFORE' in p['text']), None)
        o_section = next((p for p in o['paras'] if '匿名データの提供' in p['text']), None)
        o_after = next((p for p in o['paras'] if 'AFTER' in p['text']), None)

        if not (w_section and o_section):
            print(f'{label}: missing section header')
            continue

        w_sec_gap = w_section['y'] - w_before['y'] if w_before else 0
        word_oxi_diff = o_section['y'] - w_section['y']
        print(f'{label:<35} | W={w_before["y"] if w_before else "?":>5}/{o_before["y"] if o_before else "?":>5} | '
              f'W={w_section["y"]:>5}/O={o_section["y"]:>5} | '
              f'W={w_after["y"] if w_after else "?":>5}/{o_after["y"] if o_after else "?":>5} | '
              f'{w_sec_gap:>6.2f} | {word_oxi_diff:+.2f}')


if __name__ == '__main__':
    main()
