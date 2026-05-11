"""Day 33 part 48 — Cumulative trajectory Word vs Oxi for de6e.

R6.3 found Oxi body residual +869pt over Word body residual, despite
per-paragraph spacing matching ±0.4pt. The over-pump accumulates somewhere
in the trajectory.

Approach:
- Word: walk d.Paragraphs in order, capture (i, page, y, in_table, fs, text)
- Oxi: walk layout JSON, identify each paragraph block by text aggregation,
  capture each block's first text element y + first chars of text
- Match: for each Word body paragraph, find Oxi block with text prefix match
- Compute cumulative absolute_y at each matched body paragraph
- Identify where Word abs_y vs Oxi abs_y starts to diverge

Output: per-matched-body-paragraph CSV showing Word abs_y, Oxi abs_y, diff,
and what's between this and previous match (tables / empty paragraphs).
"""
from __future__ import annotations
import os, sys, json, subprocess, re
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc
import glob

DOCX = glob.glob('tools/golden-test/documents/docx/de6e32b5960b*')[0]
RENDERER = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
PAGE_H = 841.95
WD_VPOS = 6
WD_PAGE = 3
WD_IN_TABLE = 12
MATCH_PREFIX_LEN = 6


def abs_y(pg, y):
    if pg is None or pg < 1 or y is None or y < 0:
        return None
    return (pg - 1) * PAGE_H + y


def normalize(s):
    if not s: return ''
    s = s.replace('　', ' ').replace('\r', '').replace('\x07', '').strip()
    return re.sub(r'\s+', ' ', s)


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
            paras.append({
                'i': i, 'pg': pg, 'y': y, 'in_table': in_t,
                'text': text[:60],
            })
    finally:
        d.Close(False)
        word.Quit()
    return paras


def extract_oxi_blocks(layout):
    """Group Oxi text elements by para_idx; return list of blocks in y-order."""
    blocks = {}
    for page in layout.get('pages', []):
        pg = page.get('page')
        for el in page.get('elements', []):
            if el.get('type') != 'text': continue
            pi = el.get('para_idx')
            if pi is None:
                # Cluster by (page, y-line)
                pi = ('y', pg, round(el.get('y', 0) * 2) / 2)
            slot = blocks.setdefault(pi, {
                'para_idx': pi if not isinstance(pi, tuple) else None,
                'pg': pg,
                'y': el.get('y', 9999),
                'x': el.get('x', 9999),
                'texts': [],
            })
            slot['pg'] = min(slot['pg'], pg)
            if el.get('y', 9999) < slot['y']:
                slot['y'] = el['y']
                slot['x'] = el['x']
            slot['texts'].append((el['x'], el.get('text', '')))
    out = []
    for k, v in blocks.items():
        v['texts'].sort(key=lambda t: t[0])
        full = ''.join(t for _, t in v['texts'])
        v['text'] = full[:60]
        out.append(v)
    out.sort(key=lambda b: (b['pg'], b['y']))
    return out


def main():
    print('Measuring Word...')
    word_paras = measure_word()
    body = [p for p in word_paras if not p['in_table']]
    print(f'  total={len(word_paras)}, body={len(body)}')

    print('Rendering Oxi...')
    out_layout = os.path.join(r'C:\tmp', 'de6e_traj_layout.json')
    cmd = [RENDERER, os.path.abspath(DOCX), os.path.join(r'C:\tmp', 'de6e_traj'),
           f'--dump-layout={out_layout}']
    subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    oxi_blocks = extract_oxi_blocks(layout)
    print(f'  Oxi blocks: {len(oxi_blocks)}')

    # Match each Word body paragraph to an Oxi block by text prefix
    matches = []
    used_oxi_idx = set()
    for wp in body:
        wt = normalize(wp['text'])
        if len(wt) < MATCH_PREFIX_LEN:
            matches.append({**wp, 'oxi_pg': None, 'oxi_y': None, 'oxi_text': None})
            continue
        prefix = wt[:MATCH_PREFIX_LEN]
        best = None
        for j, ob in enumerate(oxi_blocks):
            if j in used_oxi_idx: continue
            ot = normalize(ob['text'])
            if ot.startswith(prefix) or (len(ot) >= 3 and wt.startswith(ot[:6])):
                best = (j, ob)
                break
        if best:
            j, ob = best
            used_oxi_idx.add(j)
            matches.append({**wp, 'oxi_pg': ob['pg'], 'oxi_y': round(ob['y'], 2),
                          'oxi_text': ob['text'][:30]})
        else:
            matches.append({**wp, 'oxi_pg': None, 'oxi_y': None, 'oxi_text': None})

    print(f'\nMatched: {sum(1 for m in matches if m["oxi_pg"] is not None)}/{len(matches)} body paragraphs')
    print(f'\n{"i":>4} {"w_pg":>4} {"w_y":>7} {"w_ay":>8} {"o_pg":>4} {"o_y":>7} {"o_ay":>8} {"diff":>7}  text')
    prev_w_ay = None
    prev_o_ay = None
    cumul_diff = 0.0
    for m in matches:
        w_ay = abs_y(m['pg'], m['y'])
        o_ay = abs_y(m['oxi_pg'], m['oxi_y']) if m['oxi_pg'] else None
        diff = (o_ay - w_ay) if (w_ay is not None and o_ay is not None) else None
        if diff is not None:
            cumul_diff = diff
        w_ay_str = f'{w_ay:>8.2f}' if w_ay else '       ?'
        o_ay_str = f'{o_ay:>8.2f}' if o_ay else '       ?'
        diff_str = f'{diff:+7.2f}' if diff is not None else '       '
        opg = m['oxi_pg'] if m['oxi_pg'] else '?'
        oy = m['oxi_y'] if m['oxi_y'] else '?'
        print(f'{m["i"]:>4} {m["pg"]:>4} {m["y"]:>7} {w_ay_str} {opg:>4} {oy:>7} {o_ay_str} {diff_str}  {m["text"][:35]!r}')

    print(f'\nFinal cumulative diff: {cumul_diff:+.2f}pt (Oxi - Word at last matched body paragraph)')


if __name__ == '__main__':
    main()
