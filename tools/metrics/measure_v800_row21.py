"""S155: Measure V800 row21 drift isolation."""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'a1d6_row21_isolate')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

VARIANTS = [
    'V800a_overflow_valign_trH_atLeast',
    'V800b_overflow_valign_top',
    'V800c_overflow_no_trH',
    'V800d_overflow_trH_exact',
    'V800e_single_para_trH',
    'V800f_overflow_no_sb',
    'V800g_1para_trH_atLeast',
    'V800h_2cell_a1d6_exact',
    'V800i_1para_no_trH',
]


def measure(docx_path):
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(docx_path, ReadOnly=True)
    w_y = w_pg = None
    try:
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            if '取り扱う' in (p.Range.Text or ''):
                rng = p.Range
                cr = d.Range(rng.Start, rng.Start)
                w_y = round(cr.Information(6), 2)
                w_pg = int(cr.Information(3))
                break
    finally:
        d.Close(False)
        word.Quit()
    # Oxi
    label = os.path.splitext(os.path.basename(docx_path))[0]
    out = os.path.join(r'C:\tmp', f'{label}.json')
    r = subprocess.run([RENDERER, docx_path, os.path.join(r'C:\tmp', label), f'--dump-layout={out}'],
                       capture_output=True, text=True)
    o_y = o_pg = None
    if r.returncode == 0:
        layout = json.load(open(out, encoding='utf-8'))
        for pi, page in enumerate(layout['pages']):
            for el in page['elements']:
                if el.get('type') == 'text' and '取り扱う' in el.get('text', ''):
                    o_y = round(el['y'], 2)
                    o_pg = pi + 1
                    break
            if o_y is not None: break
    return w_y, w_pg, o_y, o_pg


def main():
    print(f'{"Variant":<40} | Word pg/y | Oxi pg/y | drift')
    print('-' * 90)
    for label in VARIANTS:
        docx = os.path.join(REPRO_DIR, f'{label}.docx')
        if not os.path.exists(docx):
            print(f'{label}: NOT FOUND')
            continue
        try:
            w_y, w_pg, o_y, o_pg = measure(docx)
        except Exception as e:
            print(f'{label}: ERROR: {e}')
            continue
        if w_y is None or o_y is None:
            print(f'{label}: missing')
            continue
        drift = o_y - w_y
        print(f'{label:<40} | pg{w_pg} y={w_y:>6} | pg{o_pg} y={o_y:>6} | {drift:+.2f}')


if __name__ == '__main__':
    main()
