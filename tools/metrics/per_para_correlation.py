"""Day 32 part 10 — Per-paragraph dy ↔ attribute correlation.

Day 32 part 9 closed single-attribute detector path. Now collecting
per-paragraph (Word y, Oxi y, dy, attributes) for ALL paragraphs in
4 Class A + 3 preserve docs, output CSV for multi-attribute analysis.

Attributes per paragraph:
- font_size (Word COM Range.Font.Size first run)
- line_spacing_rule (LineSpacingRule)
- line_spacing_val (LineSpacing)
- snap_to_grid (SnapToGrid Format)
- style_name (Style.NameLocal)
- in_table (Range.Information(12))
- text_alignment_para (Format.TextAlignment, vertical)
- has_indent (LeftIndent / FirstLineIndent != 0)
- is_empty (Text.strip() == '')
- y_word (cr_start.Information(6))
- y_oxi (matched layout para_idx first y)
- dy = oxi_y - word_y (positive = Oxi later)

Output: pipeline_data/per_para_correlation_<doc_id>.csv
"""
from __future__ import annotations
import os, sys, json, subprocess, re, csv
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX_DIR = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx')
RENDERER = os.path.abspath(os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe'))
OUT_DIR = os.path.join(REPO, 'pipeline_data')
PAGE_HEIGHT = 841.95
MATCH_PREFIX_LEN = 8


def find_docx(doc_id):
    for f in os.listdir(DOCX_DIR):
        if f.startswith(doc_id) and f.endswith('.docx'):
            return os.path.join(DOCX_DIR, f)
    return None


def normalize(s):
    if not s:
        return ''
    s = s.replace('　', ' ').replace('\r', '').replace('\x07', '').strip()
    s = re.sub(r'\s+', ' ', s)
    return s


def render_oxi(docx):
    label = os.path.splitext(os.path.basename(docx))[0]
    out_layout = os.path.join(r'C:\tmp', f'{label}_v2_layout.json')
    if not os.path.exists(out_layout):
        cmd = [RENDERER, docx, os.path.join(r'C:\tmp', f'{label}_v2'), f'--dump-layout={out_layout}']
        subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    by_loc = {}
    for page in layout.get('pages', []):
        pg = page.get('page')
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            y = round(el.get('y', 0), 1)
            pi = el.get('para_idx')
            if pi is None:
                pi = -1
            key = (pg, y, pi)
            by_loc.setdefault(key, []).append(el.get('text', ''))
    out = []
    for (pg, y, pi), texts in sorted(by_loc.items()):
        full = ''.join(texts)
        out.append({'page': pg, 'y': y, 'text': full[:80], 'para_idx': pi})
    return out


def measure_word_full(docx):
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
    paras = []
    try:
        n = d.Paragraphs.Count
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr = d.Range(r.Start, r.Start)
            text = (r.Text or '').strip()
            try: fs = r.Font.Size
            except: fs = -1
            try: lh_rule = p.Format.LineSpacingRule
            except: lh_rule = -1
            try: lh_val = p.Format.LineSpacing
            except: lh_val = -1
            try: snap = p.Format.SnapToGrid
            except: snap = -1
            try: style_name = str(p.Style.NameLocal)
            except: style_name = '?'
            try: in_table = bool(r.Information(12))
            except: in_table = False
            try: ta = p.Format.TextAlignment
            except: ta = -1
            try: li = p.Format.LeftIndent
            except: li = 0
            try: fli = p.Format.FirstLineIndent
            except: fli = 0
            paras.append({
                'i': i,
                'text': text[:50],
                'is_empty': len(text) == 0,
                'page': int(cr.Information(3)),
                'y': round(cr.Information(6), 2),
                'fs': fs,
                'lh_rule': lh_rule,
                'lh_val': round(lh_val, 2) if lh_val else lh_val,
                'snap': snap,
                'style_name': style_name,
                'in_table': in_table,
                'text_align': ta,
                'left_indent': round(li, 1),
                'first_line_indent': round(fli, 1),
            })
    finally:
        d.Close(False)
        word.Quit()
    return paras


def match_text(word_paras, oxi_lines):
    matches = {}  # word_i -> match dict
    used_oxi = set()
    for w in word_paras:
        wt = normalize(w['text'])
        if not wt or len(wt) < MATCH_PREFIX_LEN:
            continue
        prefix = wt[:MATCH_PREFIX_LEN]
        for j, o in enumerate(oxi_lines):
            if j in used_oxi:
                continue
            ot = normalize(o['text'])
            if ot.startswith(prefix):
                w_abs = (w['page'] - 1) * PAGE_HEIGHT + w['y']
                o_abs = (o['page'] - 1) * PAGE_HEIGHT + o['y']
                matches[w['i']] = {
                    'oxi_pg': o['page'], 'oxi_y': o['y'],
                    'oxi_para_idx': o['para_idx'],
                    'dy_abs': round(o_abs - w_abs, 2),
                }
                used_oxi.add(j)
                break
    return matches


def process(doc_id, out_dir):
    docx = find_docx(doc_id)
    if not docx:
        print(f'  {doc_id}: NOT FOUND')
        return
    print(f'  {doc_id}: rendering Oxi + measuring Word...')
    oxi_lines = render_oxi(docx)
    word_paras = measure_word_full(docx)
    matches = match_text(word_paras, oxi_lines)

    rows = []
    for w in word_paras:
        m = matches.get(w['i'])
        row = {
            'doc_id': doc_id,
            'word_i': w['i'],
            'word_pg': w['page'],
            'word_y': w['y'],
            'fs': w['fs'],
            'lh_rule': w['lh_rule'],
            'lh_val': w['lh_val'],
            'snap': w['snap'],
            'style_name': w['style_name'],
            'in_table': w['in_table'],
            'text_align': w['text_align'],
            'left_indent': w['left_indent'],
            'first_line_indent': w['first_line_indent'],
            'is_empty': w['is_empty'],
            'oxi_pg': m['oxi_pg'] if m else '',
            'oxi_y': m['oxi_y'] if m else '',
            'oxi_para_idx': m['oxi_para_idx'] if m else '',
            'dy_abs': m['dy_abs'] if m else '',
            'matched': bool(m),
            'text': w['text'],
        }
        rows.append(row)

    out_path = os.path.join(out_dir, f'per_para_correlation_{doc_id}.csv')
    with open(out_path, 'w', encoding='utf-8', newline='') as f:
        if rows:
            writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            writer.writeheader()
            writer.writerows(rows)
    n_matched = sum(1 for r in rows if r['matched'])
    n_drift = sum(1 for r in rows if r['matched'] and abs(float(r['dy_abs'])) >= 0.5)
    print(f'    wrote {out_path}: {len(rows)} paras, {n_matched} matched, {n_drift} drift>=0.5pt')


def main():
    out_dir = OUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    if len(sys.argv) > 1 and sys.argv[1] == 'preserve':
        print('=== Preserve docs (no drift) ===')
        for d in ['e3c545fac7a7', '0e7af1ae8f21', 'cb8be715d839']:
            process(d, out_dir)
    else:
        print('=== Class A docs (drift) ===')
        for d in ['bd90b00ab7a7', 'de6e32b5960b', 'db9ca18368cd', 'd77a58485f16']:
            process(d, out_dir)
        print('\n=== Preserve docs (no drift) ===')
        for d in ['e3c545fac7a7', '0e7af1ae8f21', 'cb8be715d839']:
            process(d, out_dir)


if __name__ == '__main__':
    main()
