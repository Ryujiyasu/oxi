"""Day 32 part 6 — Class A2 delta-drift localizer.

Day 32 part 5 confirmed cumulative drift in bd90b00 (~0.3pt/para after
tables, total +14pt at pi=77). Per-paragraph dy alone hides WHERE the
drift jumps. This tool computes delta-drift = dy_curr - dy_prev so
paragraph-level spikes stand out from steady accumulation.

Output table: word_i, ddy (current - previous), running dy, sample.
Spikes (|ddy| > 1pt) flag specific source paragraphs.
"""
from __future__ import annotations
import os, sys, json, subprocess, re
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX_DIR = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx')
RENDERER = os.path.abspath(os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe'))
PAGE_HEIGHT = 841.95
MATCH_PREFIX_LEN = 8


def find_docx(doc_id: str) -> str | None:
    for f in os.listdir(DOCX_DIR):
        if f.startswith(doc_id) and f.endswith('.docx'):
            return os.path.join(DOCX_DIR, f)
    return None


def normalize(s: str) -> str:
    if not s:
        return ''
    s = s.replace('　', ' ').replace('\r', '').replace('\x07', '').strip()
    s = re.sub(r'\s+', ' ', s)
    return s


def render_oxi(docx: str) -> list[dict]:
    label = os.path.splitext(os.path.basename(docx))[0]
    out_layout = os.path.join(r'C:\tmp', f'{label}_v2_layout.json')
    if not os.path.exists(out_layout):
        cmd = [RENDERER, docx, os.path.join(r'C:\tmp', f'{label}_v2'), f'--dump-layout={out_layout}']
        subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    by_loc: dict[tuple, list[str]] = {}
    for page in layout.get('pages', []):
        pg = page.get('page')
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            y = round(el.get('y', 0), 1)
            key = (pg, y, el.get('para_idx'))
            by_loc.setdefault(key, []).append(el.get('text', ''))
    out = []
    for (pg, y, pi), texts in sorted(by_loc.items()):
        full = ''.join(texts)
        out.append({'page': pg, 'y': y, 'text': full[:80], 'para_idx': pi})
    return out


def measure_word(docx: str) -> list[dict]:
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
            cr_start = d.Range(r.Start, r.Start)
            text = (r.Text or '').strip()
            in_table = False
            try:
                in_table = bool(r.Information(12))  # wdWithInTable
            except Exception:
                pass
            paras.append({
                'i': i,
                'text': text[:80],
                'page': int(cr_start.Information(3)),
                'y': round(cr_start.Information(6), 2),
                'in_table': in_table,
            })
    finally:
        d.Close(False)
        word.Quit()
    return paras


def match_text(word_paras, oxi_lines):
    matches = []
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
                matches.append({
                    'word_i': w['i'],
                    'word_pg': w['page'], 'word_y': w['y'],
                    'oxi_pg': o['page'], 'oxi_y': o['y'],
                    'in_table': w.get('in_table', False),
                    'dy_abs': round(o_abs - w_abs, 2),
                    'text': wt[:50],
                })
                used_oxi.add(j)
                break
    return matches


def analyze(doc_id):
    docx = find_docx(doc_id)
    if not docx:
        return
    print(f'\n=== {doc_id} delta-drift localizer ===')
    oxi_lines = render_oxi(docx)
    word_paras = measure_word(docx)
    matches = match_text(word_paras, oxi_lines)

    print(f'  matched: {len(matches)} of word {len(word_paras)} paras / oxi {len(oxi_lines)} lines')
    print(f'  {"w_i":>3} {"in_t":>4} {"dy":>7} {"ddy":>7} {"w_pg":>4} {"o_pg":>4} text')
    prev_dy = 0
    spikes = []
    table_jumps = []
    for m in matches:
        ddy = round(m['dy_abs'] - prev_dy, 2)
        flag = ''
        if abs(ddy) >= 1.0:
            flag = ' <<'
            spikes.append({'word_i': m['word_i'], 'ddy': ddy, 'in_table': m['in_table'], 'text': m['text']})
        in_t = 'T' if m['in_table'] else '-'
        if abs(ddy) >= 0.5 or abs(m['dy_abs']) >= 5:
            print(f'  {m["word_i"]:>3} {in_t:>4} {m["dy_abs"]:+7.2f} {ddy:+7.2f} {m["word_pg"]:>4} {m["oxi_pg"]:>4} {m["text"]!r}{flag}')
        prev_dy = m['dy_abs']

    print(f'\n  Spikes (|ddy|>=1pt): {len(spikes)}')
    print('  --- All spikes ---')
    for s in spikes:
        print(f'    pi={s["word_i"]} ddy={s["ddy"]:+.2f} in_table={s["in_table"]} {s["text"]!r}')

    in_table_count = sum(1 for m in matches if m['in_table'])
    out_table_count = len(matches) - in_table_count
    print(f'\n  In-table paragraphs matched: {in_table_count}')
    print(f'  Out-of-table paragraphs matched: {out_table_count}')


def main():
    docs = ['bd90b00ab7a7', 'de6e32b5960b', 'db9ca18368cd', 'd77a58485f16']
    for doc_id in docs:
        analyze(doc_id)


if __name__ == '__main__':
    main()
