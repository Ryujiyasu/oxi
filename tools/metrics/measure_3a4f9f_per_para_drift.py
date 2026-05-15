"""Per-paragraph y-delta drift analysis for 3a4f9f cascade.

Per [[session59-3a4f9f-floating-table-not-cause]] the 3a4f9f cascade
(+3 pages over 1574 paragraphs) is caused by cumulative per-paragraph
drift of ~0.08pt/para — NOT by floating tables or any discrete spec.
This tool quantifies the drift by paragraph TYPE so a future fix can
target the highest-contributing type.

Methodology (linear-y approach, robust across page breaks):
  1. Load pagination_diff matches (word_i + text + word_page + oxi_page).
  2. Load pagination_word (i → page, y, in_table) and pagination_oxi.
  3. For each match, find Word's (page, y) via word_i and Oxi's (page, y)
     by text+page lookup. Convert both to linear_y = (page-1) *
     content_h + (y - top_margin). This makes drift computable across
     page boundaries.
  4. Classify Word paragraph by type (empty, heading, bracketPair,
     numbered list item, in-table, body, etc.).
  5. Compute drift_per_para = (linear_oxi - linear_word) and its DELTA
     between consecutive paragraphs to find which types cause local
     drift growth.
  6. Aggregate per type: count, mean / median drift increment, total
     contribution.

Instrumentation only — does NOT modify oxidocs-core or change any
baseline. Phase 1 53/55 mean 0.9842 must remain unchanged.
"""
from __future__ import annotations
import os, sys, json, re, statistics
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')

REPO = r'c:\Users\ryuji\oxi-main'
DOC_ID = '3a4f9fbe1a83'
DIFF = os.path.join(REPO, 'pipeline_data', 'pagination_diff', f'{DOC_ID}.json')
WORD = os.path.join(REPO, 'pipeline_data', 'pagination_word', f'{DOC_ID}.json')
OXI = os.path.join(REPO, 'pipeline_data', 'pagination_oxi', f'{DOC_ID}.json')
OUT = os.path.join(REPO, 'pipeline_data', 'ra_manual_measurements', f'{DOC_ID}_per_para_drift.json')

# A4 3a4f9f geometry (twips → pt for the layout area):
# pgSz w=11906 h=16838 (= 595x842pt). pgMar top=1985tw=99.25pt,
# bottom=1701tw=85.05pt. content_h = 842 - 99.25 - 85.05 = 657.7pt.
PAGE_HEIGHT = 842.0
TOP_MARGIN = 99.25
BOTTOM_MARGIN = 85.05
CONTENT_H = PAGE_HEIGHT - TOP_MARGIN - BOTTOM_MARGIN  # 657.7pt


# Paragraph type classification — based on Word text. Order matters.
TYPE_PATTERNS = [
    ('empty',          re.compile(r'^[\s　]*$')),
    ('page_number',    re.compile(r'^[\s　]*\d{1,3}[\s　]*$')),
    ('chapter_kanji',  re.compile(r'^[\s　]*第[一二三四五六七八九十百\d０-９]+章')),
    ('article_kanji',  re.compile(r'^[\s　]*第[一二三四五六七八九十百\d０-９]+条')),
    ('bracket_pair',   re.compile(r'^[\s　]*[【〔]')),
    ('list_marker',    re.compile(r'^[\s　]*[・]')),
    ('numbered_paren', re.compile(r'^[\s　]*[（\(][一二三四五六七八九十\d０-９]')),
    ('numbered_kanji', re.compile(r'^[\s　]*[一二三四五六七八九十][\s　]')),
    ('numbered_arabic',re.compile(r'^[\s　]*[\d０-９]+[\s　]')),
    ('numbered_circle',re.compile(r'^[\s　]*[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮]')),
]


def classify(text: str, in_table: bool) -> str:
    t = (text or '')
    base = 'body'
    for label, pat in TYPE_PATTERNS:
        if pat.search(t):
            base = label
            break
    if in_table:
        return f'{base}_intable'
    return base


def linear_y(page: int | None, y: float | None) -> float | None:
    if page is None or y is None or page < 1:
        return None
    return (page - 1) * CONTENT_H + (y - TOP_MARGIN)


def main():
    with open(DIFF, encoding='utf-8') as f:
        diff = json.load(f)
    with open(WORD, encoding='utf-8') as f:
        word = json.load(f)
    with open(OXI, encoding='utf-8') as f:
        oxi = json.load(f)

    word_by_i = {p['i']: p for p in word['paragraphs']}

    oxi_by_page_text = defaultdict(list)
    for page_num, entries in oxi['pages'].items():
        pg = int(page_num)
        for e in entries:
            t = e.get('text') or ''
            oxi_by_page_text[(pg, t)].append({
                'y': e.get('y'),
                'para_idx': e.get('para_idx'),
            })

    # Walk matches; for each, compute linear_y on both sides.
    paragraphs = []
    skipped = defaultdict(int)
    for seq_idx, m in enumerate(diff['matches']):
        wi = m.get('word_i')
        wp = m.get('word_page')
        op = m.get('oxi_page')
        text = m.get('text') or ''
        delta_p = m.get('page_delta')

        wpara = word_by_i.get(wi)
        if not wpara:
            skipped['no_word_para'] += 1
            continue
        in_table = bool(wpara.get('in_table', False))
        word_y = wpara.get('y')
        word_lin = linear_y(wp, word_y)

        # Find Oxi y by (oxi_page, text). For empty/repeating texts, use
        # CLOSEST y to expected linear position.
        oxi_lin = None
        oxi_y = None
        if op is not None:
            cands = oxi_by_page_text.get((op, text), [])
            if len(cands) == 1:
                oxi_y = cands[0]['y']
                oxi_lin = linear_y(op, oxi_y)
            elif len(cands) > 1:
                # Pick the candidate with linear_y closest to word_lin
                # (page-shifted) if available; else first.
                if word_lin is not None:
                    expected_oxi_lin = word_lin  # naive: same linear position
                    best = min(cands, key=lambda c: abs(
                        (linear_y(op, c['y']) or 0) - expected_oxi_lin))
                    oxi_y = best['y']
                    oxi_lin = linear_y(op, oxi_y)
                    skipped['multi_oxi_resolved_by_distance'] += 1
                else:
                    oxi_y = cands[0]['y']
                    oxi_lin = linear_y(op, oxi_y)
                    skipped['multi_oxi_first'] += 1
            else:
                skipped['no_oxi_text_match'] += 1

        ptype = classify(text, in_table)
        if word_lin is not None and oxi_lin is not None:
            drift = round(oxi_lin - word_lin, 3)
        else:
            drift = None

        paragraphs.append({
            'seq_idx': seq_idx,
            'word_i': wi,
            'word_page': wp,
            'word_y': word_y,
            'oxi_page': op,
            'oxi_y': oxi_y,
            'word_linear': round(word_lin, 2) if word_lin is not None else None,
            'oxi_linear': round(oxi_lin, 2) if oxi_lin is not None else None,
            'drift': drift,
            'page_delta': delta_p,
            'type': ptype,
            'text_preview': text[:30],
        })

    # Filter to records with valid drift
    valid = [p for p in paragraphs if p['drift'] is not None]
    print(f'=== 3a4f9f per-paragraph drift (linear-y approach) ===')
    print(f'Total matches: {len(diff["matches"])}')
    print(f'Records with valid drift: {len(valid)}')
    print(f'Skipped: {dict(skipped)}')

    if not valid:
        print('No valid records — aborting.')
        return

    # Drift increment between consecutive matched paragraphs (sorted by word_i)
    valid_sorted = sorted(valid, key=lambda p: p['word_i'])
    type_increments = defaultdict(list)
    prev_drift = None
    prev_type = None
    for p in valid_sorted:
        if prev_drift is not None:
            inc = round(p['drift'] - prev_drift, 4)
            type_increments[p['type']].append(inc)
        prev_drift = p['drift']
        prev_type = p['type']

    # Per-type stats: mean/median absolute drift AND mean increment-into-this-type
    type_summary = []
    by_type_records = defaultdict(list)
    for p in valid:
        by_type_records[p['type']].append(p['drift'])
    for t, drifts in sorted(by_type_records.items(), key=lambda kv: -len(kv[1])):
        incs = type_increments.get(t, [])
        type_summary.append({
            'type': t,
            'count': len(drifts),
            'mean_abs_drift': round(statistics.fmean(drifts), 3),
            'median_abs_drift': round(statistics.median(drifts), 3),
            'mean_increment_into': round(statistics.fmean(incs), 4) if incs else None,
            'median_increment_into': round(statistics.median(incs), 4) if incs else None,
            'sum_increment_into': round(sum(incs), 2) if incs else None,
        })

    print()
    print(f'{"type":>22}  {"count":>5}  {"mean_drift":>11}  {"med_drift":>10}  {"mean_inc":>9}  {"med_inc":>9}  {"sum_inc":>9}')
    for s in type_summary:
        def fmt(v, w, p=4):
            return ('----'.rjust(w)) if v is None else f'{v:>+{w}.{p}f}'
        print(f'{s["type"]:>22}  {s["count"]:>5}  {fmt(s["mean_abs_drift"],11,3)}  '
              f'{fmt(s["median_abs_drift"],10,3)}  {fmt(s["mean_increment_into"],9)}  '
              f'{fmt(s["median_increment_into"],9)}  {fmt(s["sum_increment_into"],9,2)}')

    # Trace: drift at every 100 paragraphs (sorted by word_i)
    print()
    print('=== Cumulative drift trace (every 100 word_i) ===')
    print(f'{"seq":>5}  {"word_i":>5}  {"wp":>3}  {"op":>3}  {"Δp":>4}  {"drift_pt":>10}  type')
    step = max(1, len(valid_sorted) // 20)
    for i in range(0, len(valid_sorted), step):
        p = valid_sorted[i]
        print(f'{p["seq_idx"]:>5}  {p["word_i"]:>5}  {p["word_page"]:>3}  '
              f'{p["oxi_page"]:>3}  {p["page_delta"]:>+4d}  {p["drift"]:>+10.2f}  {p["type"]}')
    # Final
    p = valid_sorted[-1]
    print(f'{p["seq_idx"]:>5}  {p["word_i"]:>5}  {p["word_page"]:>3}  '
          f'{p["oxi_page"]:>3}  {p["page_delta"]:>+4d}  {p["drift"]:>+10.2f}  {p["type"]}  (final)')

    # Save
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    payload = {
        'doc_id': DOC_ID,
        'word_filename': diff.get('word_filename'),
        'page_geometry': {
            'page_height': PAGE_HEIGHT, 'top_margin': TOP_MARGIN,
            'bottom_margin': BOTTOM_MARGIN, 'content_h': CONTENT_H,
        },
        'n_matches': len(diff['matches']),
        'n_valid': len(valid),
        'skipped': dict(skipped),
        'type_summary': type_summary,
        'paragraphs': paragraphs,
    }
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f'\nSaved to {OUT}')


if __name__ == '__main__':
    main()
