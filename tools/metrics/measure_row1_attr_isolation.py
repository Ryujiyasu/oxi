"""Day 33 part 16 — Word + Oxi row height comparison for row1_attr_isolation repros."""
from __future__ import annotations
import os, sys, json, glob, subprocess
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'row1_attr_isolation')
OUT = os.path.join(REPO, 'pipeline_data', 'row1_attr_isolation.json')
GDI_BIN = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')


def measure_word(word, docx_path: str) -> dict:
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    out = {}
    try:
        row = d.Tables(1).Rows(1)
        cr_row = d.Range(row.Range.Start, row.Range.Start)
        out['row_y'] = cr_row.Information(6)
        # Anchor (paragraph after table)
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            if '下のアンカー' in (p.Range.Text or ''):
                cr = d.Range(p.Range.Start, p.Range.Start)
                out['anchor_y'] = cr.Information(6)
                break
        # First text para in cell
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            if '提供申出者' in (p.Range.Text or ''):
                cr = d.Range(p.Range.Start, p.Range.Start)
                out['cell_text_y'] = cr.Information(6)
                break
        if 'anchor_y' in out and 'row_y' in out:
            out['row_height'] = out['anchor_y'] - out['row_y']
    finally:
        d.Close(SaveChanges=False)
    return out


def measure_oxi(docx_path: str, label: str) -> dict:
    layout_json = os.path.join(r'C:\tmp', f'{label}_oxi_r1a.json')
    out_prefix = os.path.join(r'C:\tmp', label + '_oxi_r1a')
    res = subprocess.run([
        GDI_BIN, docx_path, out_prefix, '96', f'--dump-layout={layout_json}'
    ], capture_output=True, text=True, timeout=60)
    if res.returncode != 0:
        return {'error': res.stderr[:200]}
    layout = json.load(open(layout_json, encoding='utf-8'))
    out = {}
    cell_text_y = None
    anchor_y = None
    top_border_y = None
    bottom_border_y = None
    for page in layout.get('pages', []):
        if page.get('page') != 1: continue
        for el in page.get('elements', []):
            if el.get('type') == 'text':
                txt = el.get('text', '')
                if '提供申出者' in txt and cell_text_y is None:
                    cell_text_y = el.get('y')
                # Anchor text 「下のアンカー」 may be split per char; first char '下' marks anchor.
                if anchor_y is None and (txt == '下' or '下のアンカー' in txt or txt.startswith('下')):
                    anchor_y = el.get('y')
            elif el.get('type') == 'border' and el.get('w', 0) > 100 and el.get('h', 1) < 1:
                y = el.get('y')
                if top_border_y is None or y < top_border_y:
                    top_border_y = y
                if bottom_border_y is None or y > bottom_border_y:
                    bottom_border_y = y
    out['cell_text_y'] = cell_text_y
    out['anchor_y'] = anchor_y
    out['top_border_y'] = top_border_y
    out['bottom_border_y'] = bottom_border_y
    if top_border_y is not None and bottom_border_y is not None:
        out['oxi_row_h_borders'] = round(bottom_border_y - top_border_y, 3)
    if cell_text_y is not None and anchor_y is not None:
        out['cell_text_to_anchor'] = anchor_y - cell_text_y
    return out


def main():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    docs = sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx')))
    results = []
    try:
        for fp in docs:
            label = os.path.basename(fp).replace('.docx', '')
            wm = measure_word(word, fp)
            om = measure_oxi(fp, label)
            r = {'label': label, 'word': wm, 'oxi': om}
            # Compare row heights
            wh = wm.get('row_height')
            wt = wm.get('cell_text_y')
            wa = wm.get('anchor_y')
            ot = om.get('cell_text_y')
            oa = om.get('anchor_y')
            if wt and wa and ot and oa:
                w_advance = wa - wt  # cell text to anchor (excl. cell padding top of cell)
                o_advance = oa - ot
                r['word_text_to_anchor'] = round(w_advance, 3)
                r['oxi_text_to_anchor'] = round(o_advance, 3)
                r['diff'] = round(o_advance - w_advance, 3)
            results.append(r)
            wh_s = f'{wh:.2f}' if wh else 'ERR'
            wta = r.get('word_text_to_anchor', '?')
            ota = r.get('oxi_text_to_anchor', '?')
            df = r.get('diff', '?')
            print(f'{label:30s}  word_row_h={wh_s:>7}  word_t2a={wta!s:>7}  oxi_t2a={ota!s:>7}  diff={df!s:>7}')
    finally:
        word.Quit()
    json.dump(results, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f'\nSaved: {OUT}')

if __name__ == '__main__':
    main()
