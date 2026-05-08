"""Measure per-paragraph line count for db9ca full document.

Day 31 part 13 found that Oxi's cursor at i=19 is +18pt below Word's,
implying one of paragraphs 1-18 wraps to 1 extra line in Oxi vs Word's
full-doc rendering. DW_V113 (isolated para 18) matched 2L/2L, so the
divergence requires preceding context.

Strategy:
  - Word: use paragraph end_y - start_y to derive line count.
    For multi-line paragraph spanning [start_y, end_y], lines = round(span / line_height) + 1
    when on same page.
  - Oxi: --dump-layout JSON; count DISTINCT y values per para_idx within page.
  - Cross-match by paragraph index, compare line counts, identify divergence.
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx',
                    'db9ca18368cd_20241122_resource_open_data_01.docx')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TMP = r'C:\tmp'


def measure_word(docx_path: str) -> list[dict]:
    """Per-paragraph: i, page, start_y, end_y, char_count, line_height."""
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
            # Range end-1 to avoid trailing paragraph mark on next page
            end_pos = max(r.Start, r.End - 1)
            cr_end = d.Range(end_pos, end_pos)
            text = (r.Text or '').rstrip()
            paras.append({
                'i': i,
                'text': text[:60],
                'char_count': len(text),
                'start_page': int(cr_start.Information(3)),
                'end_page': int(cr_end.Information(3)),
                'start_y': round(cr_start.Information(6), 2),
                'end_y': round(cr_end.Information(6), 2),
                'fmt_line_spacing': float(p.Format.LineSpacing) if p.Format else 0,
            })
    finally:
        d.Close(False)
        word.Quit()
    return paras


def measure_oxi(docx_path: str) -> dict:
    label = os.path.splitext(os.path.basename(docx_path))[0]
    out_prefix = os.path.join(TMP, f'{label}_lc')
    out_layout = os.path.join(TMP, f'{label}_lc_layout.json')
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
                by_para[pi] = {
                    'para_idx': pi,
                    'lines': [],
                    'pages': set(),
                }
            by_para[pi]['lines'].append({
                'y': el['y'], 'h': el['h'], 'page': page.get('page', page_idx + 1),
                'text': el.get('text', '')[:30],
            })
            by_para[pi]['pages'].add(page.get('page', page_idx + 1))
    out = {}
    for pi, p in by_para.items():
        # Distinct y values within paragraph (across all pages)
        unique_ys = sorted(set(round(line['y'], 1) for line in p['lines']))
        out[pi] = {
            'para_idx': pi,
            'n_lines': len(unique_ys),
            'first_y': min(line['y'] for line in p['lines']),
            'last_y': max(line['y'] for line in p['lines']),
            'pages': sorted(p['pages']),
        }
    return out


def estimate_word_lines(p: dict, default_lh: float = 18.0) -> int:
    """Estimate Word's line count from start_y, end_y, paragraph format.
    Single-page paragraphs: lines = round((end_y - start_y) / line_height) + 1.
    """
    if p['start_page'] != p['end_page']:
        return -1  # cross-page, can't estimate easily
    span = p['end_y'] - p['start_y']
    if span < 0.5:
        return 1
    # line_height: Format.LineSpacing returns line spacing if exact, else single line height
    lh = p.get('fmt_line_spacing') or default_lh
    # Word's LineSpacing for "single" returns 12 (font size in pt typically),
    # for "exact 18pt" returns 18. Use a sensible default if 0.
    if lh < 5:
        lh = default_lh
    return round(span / lh) + 1


def main():
    print('Measuring Word...')
    word_paras = measure_word(DOCX)
    print(f'  {len(word_paras)} paragraphs')
    print('Measuring Oxi...')
    oxi_paras = measure_oxi(DOCX)
    print(f'  {len(oxi_paras)} paragraphs')

    print()
    print(f'{"i":>3} {"chars":>5} {"w_pg":>4} {"w_y":>6} {"end_y":>6} {"w_lines":>7} {"o_lines":>7} {"diff":>5}  text')
    divergent = []
    for wp in word_paras:
        oxi_p = oxi_paras.get(wp['i'] - 1)  # Oxi para_idx is 0-indexed
        w_lines = estimate_word_lines(wp)
        o_lines = oxi_p['n_lines'] if oxi_p else 0
        diff = o_lines - w_lines if w_lines >= 0 else 'X'
        marker = ' <--' if isinstance(diff, int) and diff != 0 else ''
        print(f'{wp["i"]:>3} {wp["char_count"]:>5} {wp["start_page"]:>4} {wp["start_y"]:>6.1f} {wp["end_y"]:>6.1f} {w_lines:>7} {o_lines:>7} {str(diff):>5}  {wp["text"]!r}{marker}')
        if isinstance(diff, int) and diff != 0:
            divergent.append((wp['i'], w_lines, o_lines, wp['text']))

    print()
    print(f'Divergent paragraphs: {len(divergent)}')
    for i, w, o, t in divergent:
        print(f'  i={i} word={w} oxi={o} diff={o-w:+d}  {t!r}')

    out_path = os.path.join(REPO, 'pipeline_data', 'db9ca_line_counts.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump({
            'word': word_paras,
            'oxi': {str(k): v for k, v in oxi_paras.items()},
            'divergent': divergent,
        }, f, ensure_ascii=False, indent=2, default=list)
    print(f'\nWrote {out_path}')


if __name__ == '__main__':
    main()
