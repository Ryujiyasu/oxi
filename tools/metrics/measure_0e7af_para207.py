"""Measure Oxi vs Word line count for 0e7af paragraph 207 minimal repros.

Day 31 part 30: verify wrap-width hypothesis.
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', '0e7af_class_b')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TMP = r'C:\tmp'

VARIANTS = ['CB_V200_para207_only', 'CB_V201_paras_205_to_209', 'CB_V202_paras_200_to_215']


def measure_word(docx_path: str) -> list[dict]:
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(docx_path, ReadOnly=True)
    paras = []
    try:
        n = d.Paragraphs.Count
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr_start = d.Range(r.Start, r.Start)
            end_pos = max(r.Start, r.End - 1)
            cr_end = d.Range(end_pos, end_pos)
            text = (r.Text or '').rstrip('\r\n\x07')
            paras.append({
                'i': i,
                'text': text[:80],
                'len': len(text),
                'start_y': round(cr_start.Information(6), 2),
                'end_y': round(cr_end.Information(6), 2),
                'start_page': int(cr_start.Information(3)),
                'end_page': int(cr_end.Information(3)),
            })
    finally:
        d.Close(False)
        word.Quit()
    return paras


def measure_oxi(docx_path: str) -> dict:
    label = os.path.splitext(os.path.basename(docx_path))[0]
    out_prefix = os.path.join(TMP, f'{label}')
    out_layout = os.path.join(TMP, f'{label}_layout.json')
    cmd = [RENDERER, docx_path, out_prefix, f'--dump-layout={out_layout}']
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        return {}
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    pages = layout.get('pages', [])
    by_para = {}
    for page_idx, page in enumerate(pages):
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            pi = el.get('para_idx')
            if pi is None:
                continue
            if pi not in by_para:
                by_para[pi] = {'lines': [], 'pages': set()}
            by_para[pi]['lines'].append({'y': el['y'], 'text': el.get('text','')[:30]})
            by_para[pi]['pages'].add(page.get('page', page_idx + 1))
    out = {}
    for pi, p in by_para.items():
        unique_ys = sorted(set(round(line['y'], 1) for line in p['lines']))
        out[pi] = {
            'para_idx': pi,
            'n_lines': len(unique_ys),
            'first_y': min(line['y'] for line in p['lines']),
            'last_y': max(line['y'] for line in p['lines']),
            'pages': sorted(p['pages']),
        }
    return out


def main():
    for label in VARIANTS:
        docx = os.path.join(REPRO_DIR, f'{label}.docx')
        if not os.path.exists(docx):
            continue
        print(f'=== {label} ===')
        try:
            wp = measure_word(docx)
        except Exception as e:
            print(f'  Word ERROR: {e}')
            continue
        op = measure_oxi(docx)
        # Identify the paragraph 207 equivalent (has '受託者が第２３条又は第２４条' text)
        target_text = '受託者が第２３条又は第２４条'
        for w in wp:
            text = w.get('text', '')
            if target_text in text:
                # Estimate Word lines from start_y/end_y span (assuming 18pt line height for 9pt × 2)
                lh_estimate = 13.0  # 9pt font ~ 12-14pt lh in tight settings
                same_page = w['start_page'] == w['end_page']
                if same_page:
                    span = w['end_y'] - w['start_y']
                    word_lines = round(span / lh_estimate) + 1
                else:
                    word_lines = -1
                print(f'  Word i={w["i"]} pg={w["start_page"]}->{w["end_page"]} y={w["start_y"]}->{w["end_y"]} (~{word_lines} lines @ lh=13) text={text[:50]!r}')
                # Match to Oxi para
                # Oxi para_idx is 0-based for first paragraph
                oxi_idx = w['i'] - 1
                if oxi_idx in op:
                    o = op[oxi_idx]
                    print(f'  Oxi para_idx={oxi_idx} n_lines={o["n_lines"]} y={o["first_y"]}->{o["last_y"]} pages={o["pages"]}')
                    print(f'  Word/Oxi line count: {word_lines} / {o["n_lines"]} → diff={word_lines - o["n_lines"]:+}')
                break
        print()


if __name__ == '__main__':
    main()
