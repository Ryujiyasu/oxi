"""Cross-page content-alignment diff tool.

Given:
- Oxi --dump-layout JSON
- Word COM paragraph data (iterated as {idx, page, y, text})

Output per-Word-paragraph:
  word_idx | word_page | word_y | oxi_page | oxi_y | dy | dpage | text[:30]

Matches by text substring. Useful for identifying page-boundary drift.

Usage:
    python tools/metrics/cross_page_diff.py <oxi_dump.json> <word_paras.json>
"""
import json
import sys

def load_oxi(path):
    """Return list of Oxi paragraphs with {page, para_idx, first_y, text_str}."""
    with open(path, encoding='utf-8') as f:
        d = json.load(f)
    oxi_paras = []
    for p in d['pages']:
        page = p['page']
        by_para = {}
        for e in p['elements']:
            if e.get('type') != 'text': continue
            pi = e.get('para_idx')
            if pi is None: continue
            by_para.setdefault(pi, []).append(e)
        for pi, elems in by_para.items():
            elems.sort(key=lambda e: (e['y'], e['x']))
            joined = ''.join(e.get('text', '') for e in elems)
            first_y = min(e['y'] for e in elems)
            oxi_paras.append({
                'page': page,
                'para_idx': pi,
                'first_y': first_y,
                'text_str': joined,
            })
    return oxi_paras

def find_in_oxi(oxi_paras, target_str, min_match=8):
    """Find Oxi paragraph whose text contains the target string."""
    if len(target_str) < min_match:
        return None
    key = target_str[:min_match]
    for op in oxi_paras:
        if key in op['text_str']:
            return op
    return None

def main():
    if len(sys.argv) < 3:
        print("Usage: cross_page_diff.py <oxi_dump.json> <word_paras.json>")
        sys.exit(1)

    oxi_paras = load_oxi(sys.argv[1])
    with open(sys.argv[2], encoding='utf-8') as f:
        word_paras = json.load(f)

    print(f'Oxi paragraphs: {len(oxi_paras)}')
    print(f'Word paragraphs: {len(word_paras)}')
    print()
    print(f'{"w_idx":>5} {"w_pg":>4} {"w_y":>6} | {"o_pg":>4} {"o_y":>6} | {"dpg":>4} {"dy":>6} | text')
    print('-' * 100)

    matched = 0
    page_drift = {}
    for wp in word_paras:
        if 'error' in wp: continue
        idx = wp.get('idx')
        w_page = wp.get('page')
        w_y = wp.get('y')
        # Preferred: decode sjis_hex to proper Unicode string
        if 'sjis_hex' in wp:
            try:
                target_str = bytes.fromhex(wp['sjis_hex']).decode('shift_jis', errors='ignore')[:20]
            except Exception:
                continue
        else:
            target_str = wp.get('text', '')[:20]
        if len(target_str.strip()) < 5: continue
        text = target_str

        op = find_in_oxi(oxi_paras, target_str)
        if op is None:
            print(f'{idx:>5} {w_page:>4} {w_y:>6.1f} | NOT FOUND {target_str[:30]!r}')
            continue

        o_page = op['page']
        o_y = op['first_y']
        dpg = o_page - w_page
        dy = o_y - w_y
        marker = '!' if abs(dpg) >= 1 else ' '
        print(f'{idx:>5} {w_page:>4} {w_y:>6.1f} | {o_page:>4} {o_y:>6.1f} | {dpg:>4} {dy:>+6.1f} {marker} {text[:30]!r}')
        matched += 1
        page_drift.setdefault(dpg, 0)
        page_drift[dpg] += 1

    print(f'\nMatched: {matched}')
    print(f'Page drift distribution: {sorted(page_drift.items())}')

if __name__ == '__main__':
    main()
