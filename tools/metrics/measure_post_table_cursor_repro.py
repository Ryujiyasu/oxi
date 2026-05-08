"""Measure PT_V1-V5 minimal repros: Word vs Oxi position of 备考 paragraph
after table. Goal: localize Oxi's +4pt post-table cursor over-advance.

For each variant:
1. Word COM: page setup, last cell.Range.End y, 备考 paragraph y/page.
2. Oxi: render via oxi-gdi-renderer --dump-layout, find 备考 element y/page,
   find max border y on each page (table bottom).
3. Compare and emit a one-line summary per variant.
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'post_tbl_cursor')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TMP = r'C:\tmp'

VARIANTS = ['PT_V1_simple_with_flag', 'PT_V2_with_widowctl_empty',
            'PT_V3_with_3trailing_empties', 'PT_V4_3trailing_NO_flag',
            'PT_V5_simple_NO_flag',
            'PT_V6_tight_simple_with_flag', 'PT_V7_tight_3trailing_with_flag',
            'PT_V8_tight_3trailing_NO_flag']


def measure_word(docx_path: str) -> dict:
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(docx_path, ReadOnly=True)
    out = {}
    try:
        ps = d.PageSetup
        out['pgH'] = ps.PageHeight
        out['top'] = ps.TopMargin
        out['bottom'] = ps.BottomMargin
        out['body_bottom'] = ps.PageHeight - ps.BottomMargin
        # Find 備考 paragraph
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            txt = (p.Range.Text or '').strip()
            if txt.startswith('備考'):
                rng = p.Range
                cr = d.Range(rng.Start, rng.Start)
                out['bibou_x'] = cr.Information(5)
                out['bibou_y'] = cr.Information(6)
                out['bibou_page'] = cr.Information(3)
                break
        # Last cell of last table
        if d.Tables.Count > 0:
            t = d.Tables(d.Tables.Count)
            last_cell = t.Rows(t.Rows.Count).Cells(t.Rows(t.Rows.Count).Cells.Count)
            cr = last_cell.Range
            cend = d.Range(cr.End - 1, cr.End - 1)
            out['cell_end_y'] = cend.Information(6)
            out['cell_end_page'] = cend.Information(3)
            # Per-paragraph y inside cell
            cell_paras = []
            for cp in last_cell.Range.Paragraphs:
                rr = cp.Range
                cell_paras.append({
                    'text': (rr.Text or '').strip()[:30],
                    'y': d.Range(rr.Start, rr.Start).Information(6),
                    'page': d.Range(rr.Start, rr.Start).Information(3),
                })
            out['cell_paras'] = cell_paras
    finally:
        d.Close(SaveChanges=False)
        word.Quit()
    return out


def measure_oxi(docx_path: str, label: str) -> dict:
    layout_json = os.path.join(TMP, f'{label}_layout.json')
    out_prefix = os.path.join(TMP, label)
    res = subprocess.run([
        RENDERER, docx_path, out_prefix, '96',
        f'--dump-layout={layout_json}',
    ], capture_output=True, text=True)
    if res.returncode != 0:
        print('  Oxi render error:', res.stderr[:300])
        return {}
    d = json.load(open(layout_json, encoding='utf-8'))
    out = {'n_pages': len(d['pages'])}
    # Find 備考 elements (text 備 or 考)
    bibou_locs = []
    for pi, p in enumerate(d['pages']):
        for el in p['elements']:
            if el['type'] == 'text' and el['text'] == '備':
                bibou_locs.append({'page': pi+1, 'x': el['x'], 'y': el['y'], 'h': el['h']})
    out['bibou_locs'] = bibou_locs
    # Max border y per page
    for pi, p in enumerate(d['pages']):
        borders = [e for e in p['elements'] if e['type']=='border']
        if borders:
            max_b = max(b['y']+b['h'] for b in borders)
            out[f'border_max_y_page{pi+1}'] = round(max_b, 3)
    return out


def main():
    results = []
    for label in VARIANTS:
        docx = os.path.join(REPRO_DIR, f'{label}.docx')
        if not os.path.exists(docx):
            print(f'MISSING: {docx}')
            continue
        print(f'\n=== {label} ===')
        w = measure_word(docx)
        o = measure_oxi(docx, label)
        # Compare
        word_bibou_pg = w.get('bibou_page')
        oxi_bibou_pg = o['bibou_locs'][0]['page'] if o.get('bibou_locs') else None
        print(f'Word: bibou_y={w.get("bibou_y","?"):.2f}pg={word_bibou_pg} cell_end_y={w.get("cell_end_y","?"):.2f} body_bottom={w.get("body_bottom","?"):.2f}')
        if o:
            border_max = o.get('border_max_y_page1', '?')
            print(f'Oxi : bibou_loc={o.get("bibou_locs")} border_max_y_p1={border_max} n_pages={o.get("n_pages")}')
            if w.get('cell_end_y') and isinstance(border_max, (int, float)):
                print(f'  Δ table_end (Oxi border_max - Word cell_end): {border_max - w["cell_end_y"]:+.2f}pt')
            if oxi_bibou_pg is not None and word_bibou_pg is not None:
                if oxi_bibou_pg != word_bibou_pg:
                    print(f'  *** PAGINATION DIVERGE: Word=p{word_bibou_pg} Oxi=p{oxi_bibou_pg} ***')
                else:
                    print(f'  pagination MATCH: both p{word_bibou_pg}')
        if w.get('cell_paras'):
            print('  Word cell paragraphs:')
            for cp in w['cell_paras'][-5:]:
                print(f'    pg={cp["page"]} y={cp["y"]:.2f} text={cp["text"]!r}')
        results.append({'label': label, 'word': w, 'oxi': o})
    # Save
    out_json = os.path.join(REPO, 'pipeline_data', 'post_table_cursor_repro.json')
    json.dump(results, open(out_json, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f'\nSaved: {out_json}')

if __name__ == '__main__':
    main()
