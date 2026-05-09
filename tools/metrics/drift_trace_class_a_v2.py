"""V2: text-based matching (Day 32 part 5).

Day 32 part 4 found drift_trace_class_a uses 1:1 paragraph index matching
which breaks in table contexts (Word counts each cell paragraph, Oxi
concatenates into single para_idx). This v2 uses text-based matching
similar to pagination_diff.

For each Class A doc:
1. Render Oxi at SOFT=0pt, capture (text, page, y) per text element
2. Use Word COM to capture (text, page, y) per paragraph
3. Match by text prefix (like pagination_diff does)
4. Compute per-paragraph dy for matched paragraphs
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
    """Return list of (text, page, y) for each text-element-group."""
    label = os.path.splitext(os.path.basename(docx))[0]
    out_layout = os.path.join(r'C:\tmp', f'{label}_v2_layout.json')
    cmd = [RENDERER, docx, os.path.join(r'C:\tmp', f'{label}_v2'), f'--dump-layout={out_layout}']
    subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    # Group text elements by (page, y) — typically one paragraph line
    by_loc: dict[tuple, list[str]] = {}
    for page in layout.get('pages', []):
        pg = page.get('page')
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            y = round(el.get('y', 0), 1)
            key = (pg, y, el.get('para_idx'))
            by_loc.setdefault(key, []).append(el.get('text', ''))
    # Convert to list, take first text per location
    out = []
    for (pg, y, pi), texts in sorted(by_loc.items()):
        full = ''.join(texts)
        out.append({'page': pg, 'y': y, 'text': full[:80], 'para_idx': pi})
    return out


def measure_word(docx: str) -> list[dict]:
    """Return list of (text, page, y) per Word paragraph (including table cells)."""
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
            paras.append({
                'i': i,
                'text': text[:80],
                'page': int(cr_start.Information(3)),
                'y': round(cr_start.Information(6), 2),
            })
    finally:
        d.Close(False)
        word.Quit()
    return paras


def match_text(word_paras: list[dict], oxi_lines: list[dict]) -> list[dict]:
    """Match each Word paragraph to Oxi line by text prefix."""
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
                    'oxi_para_idx': o['para_idx'],
                    'word_pg': w['page'], 'word_y': w['y'],
                    'oxi_pg': o['page'], 'oxi_y': o['y'],
                    'dy_abs': round(o_abs - w_abs, 2),
                    'text': wt[:50],
                })
                used_oxi.add(j)
                break
    return matches


def main():
    docs = ['bd90b00ab7a7', 'de6e32b5960b', 'db9ca18368cd', 'd77a58485f16']
    for doc_id in docs:
        docx = find_docx(doc_id)
        if not docx:
            continue
        print(f'\n=== {doc_id} (text-based matching) ===')
        oxi_lines = render_oxi(docx)
        word_paras = measure_word(docx)
        matches = match_text(word_paras, oxi_lines)
        print(f'  Word: {len(word_paras)} paras, Oxi: {len(oxi_lines)} text-groups, matched: {len(matches)}')
        if matches:
            # Show drift trajectory for first 20 matched paragraphs
            print(f'  {"w_i":>3} {"w_pg":>4} {"w_y":>7} {"o_pg":>4} {"o_y":>7} {"dy":>7} text')
            for m in matches[:20]:
                print(f'  {m["word_i"]:>3} {m["word_pg"]:>4} {m["word_y"]:>7.2f} {m["oxi_pg"]:>4} {m["oxi_y"]:>7.2f} {m["dy_abs"]:>+7.2f} {m["text"]!r}')
            if len(matches) > 20:
                print('  ...')
                for m in matches[-5:]:
                    print(f'  {m["word_i"]:>3} {m["word_pg"]:>4} {m["word_y"]:>7.2f} {m["oxi_pg"]:>4} {m["oxi_y"]:>7.2f} {m["dy_abs"]:>+7.2f} {m["text"]!r}')


if __name__ == '__main__':
    main()
