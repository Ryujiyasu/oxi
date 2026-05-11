"""Day 33 part 47 — Body paragraph cumulative advance Word vs Oxi for de6e.

R6.2 finding: Oxi tables -868pt vs Word, so Oxi BODY must be +868pt over-pumped
(both have 7 pages).

This tool: walk de6e paragraphs, capture per-paragraph Word y + attrs, run Oxi
GDI render and extract per-paragraph y from layout JSON via para_idx. Compute
per-paragraph advance for body paragraphs (in_table=False). Sum cumulative
advance, compare per-paragraph divergence.
"""
from __future__ import annotations
import os, sys, json, subprocess, csv
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc
import glob

DOCX = glob.glob('tools/golden-test/documents/docx/de6e32b5960b*')[0]
RENDERER = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
PAGE_H = 841.95
WD_VPOS = 6
WD_PAGE = 3
WD_IN_TABLE = 12


def abs_y(pg, y):
    if pg is None or pg < 1 or y is None or y < 0:
        return None
    return (pg - 1) * PAGE_H + y


def measure_word():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    paras = []
    try:
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr = d.Range(r.Start, r.Start)
            try: y = round(cr.Information(WD_VPOS), 2)
            except: y = -1
            try: pg = int(cr.Information(WD_PAGE))
            except: pg = -1
            try: in_t = bool(r.Information(WD_IN_TABLE))
            except: in_t = False
            try: text = (r.Text or '').replace('\r', ' ').replace('\x07', '').strip()
            except: text = ''
            try: fs = float(r.Font.Size)
            except: fs = -1
            try: sb = round(float(p.Format.SpaceBefore), 2)
            except: sb = -1
            try: sa = round(float(p.Format.SpaceAfter), 2)
            except: sa = -1
            try: lh_rule = int(p.Format.LineSpacingRule)
            except: lh_rule = -1
            try: lh_val = round(float(p.Format.LineSpacing), 2)
            except: lh_val = -1
            try: snap = int(p.Format.SnapToGrid)
            except: snap = -1
            try: style = str(p.Style.NameLocal)
            except: style = '?'
            paras.append({
                'i': i, 'pg': pg, 'y': y, 'in_table': in_t,
                'fs': fs, 'sb': sb, 'sa': sa, 'lh_rule': lh_rule, 'lh_val': lh_val,
                'snap': snap, 'style': style, 'text': text[:30],
            })
    finally:
        d.Close(False)
        word.Quit()
    return paras


def render_oxi():
    out_layout = os.path.join(r'C:\tmp', 'de6e_body_layout.json')
    cmd = [RENDERER, os.path.abspath(DOCX), os.path.join(r'C:\tmp', 'de6e_body'),
           f'--dump-layout={out_layout}']
    subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    with open(out_layout, encoding='utf-8') as f:
        return json.load(f)


def extract_oxi_paras(layout):
    """Map para_idx → (page, min_y) for first text occurrence."""
    by_idx = {}
    for page in layout.get('pages', []):
        pg = page.get('page')
        for el in page.get('elements', []):
            if el.get('type') != 'text': continue
            pi = el.get('para_idx')
            if pi is None: continue
            y = el.get('y', 0)
            existing = by_idx.get(pi)
            if existing is None or (existing[0], existing[1]) > (pg, y):
                by_idx[pi] = (pg, y)
    return by_idx


def main():
    print('Measuring Word paragraphs...')
    word_paras = measure_word()
    print(f'  {len(word_paras)} paragraphs')
    body_paras = [p for p in word_paras if not p['in_table']]
    print(f'  body: {len(body_paras)} (in_table=False)')

    print('Rendering Oxi...')
    layout = render_oxi()
    oxi_by_idx = extract_oxi_paras(layout)
    print(f'  Oxi para_idx with text: {len(oxi_by_idx)}')

    # For each Word body paragraph, find matching Oxi (by text prefix in same page region)
    # Simpler: text-prefix match to Oxi text-aggregated paragraphs
    # For now, use word_i as 0-indexed para_idx hint (Oxi assigns sequentially)
    # Note: this is approximate; may mismatch for table-context paragraphs

    # Total Word body advance: sum of consecutive body paragraph advances
    body_paras_sorted = sorted(body_paras, key=lambda p: (p['pg'], p['y']))
    if not body_paras_sorted:
        print('No body paragraphs')
        return

    print('\n=== Word body paragraph trajectory ===')
    print(f'{"i":>4} {"pg":>3} {"y":>7} {"abs_y":>8} {"adv":>7}  sb  lh sa style  text')
    total_word_body = 0.0
    prev_ay = None
    for p in body_paras_sorted[:50]:
        ay = abs_y(p['pg'], p['y'])
        adv = (ay - prev_ay) if prev_ay is not None and ay is not None else None
        adv_str = f'{adv:+7.2f}' if adv is not None else '       '
        print(f'{p["i"]:>4} {p["pg"]:>3} {p["y"]:>7} {ay if ay else "?":>8} {adv_str}  '
              f'{p["sb"]:>4} {p["lh_val"]:>4} {p["sa"]:>3} {p["style"][:6]:>6}  {p["text"]!r}')
        if adv is not None and adv > 0 and adv < 200:
            total_word_body += adv
        prev_ay = ay

    print(f'\nWord body total advance (sum positive adv 0<x<200): {total_word_body:.1f}pt')

    # Compute Oxi advance for same paragraphs (using para_idx = i - 1)
    print('\n=== Oxi body paragraph y (matched by para_idx = word_i - 1) ===')
    oxi_paras_for_body = []
    for p in body_paras_sorted:
        oxi_idx = p['i'] - 1  # Oxi typically 0-indexed
        oxi_loc = oxi_by_idx.get(oxi_idx)
        if oxi_loc:
            opg, oy = oxi_loc
            oxi_ay = abs_y(opg, oy)
            oxi_paras_for_body.append({**p, 'oxi_pg': opg, 'oxi_y': oy, 'oxi_ay': oxi_ay})

    print(f'Matched Oxi for {len(oxi_paras_for_body)} of {len(body_paras_sorted)} body paragraphs')
    print(f'\n{"i":>4} {"w_pg/y":>10} {"o_pg/y":>10} {"w_adv":>7} {"o_adv":>7} {"diff":>6}  text')
    total_oxi_body = 0.0
    prev_w_ay = prev_o_ay = None
    for p in oxi_paras_for_body[:50]:
        w_adv = (p['y'] + (p['pg']-1)*PAGE_H - prev_w_ay) if prev_w_ay is not None else None
        o_adv = (p['oxi_ay'] - prev_o_ay) if prev_o_ay is not None and p['oxi_ay'] is not None else None
        diff = (o_adv - w_adv) if (w_adv is not None and o_adv is not None) else None
        if diff is not None and -200 < diff < 200:
            pass  # ok
        diff_str = f'{diff:+6.2f}' if diff is not None else '      '
        w_adv_str = f'{w_adv:+7.2f}' if w_adv is not None else '       '
        o_adv_str = f'{o_adv:+7.2f}' if o_adv is not None else '       '
        print(f'{p["i"]:>4} {p["pg"]}/{p["y"]:>5} {p["oxi_pg"]}/{p["oxi_y"]:>5} {w_adv_str} {o_adv_str} {diff_str}  {p["text"][:20]!r}')
        if o_adv is not None and 0 < o_adv < 200:
            total_oxi_body += o_adv
        prev_w_ay = p['y'] + (p['pg']-1)*PAGE_H
        prev_o_ay = p['oxi_ay']

    print(f'\nWord body total: {total_word_body:.1f}pt')
    print(f'Oxi body total:  {total_oxi_body:.1f}pt')
    print(f'Diff (Oxi-Word): {total_oxi_body - total_word_body:+.1f}pt')


if __name__ == '__main__':
    main()
