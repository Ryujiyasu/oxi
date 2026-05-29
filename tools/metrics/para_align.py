"""S426: paragraph alignment infrastructure for per-doc drift diagnosis.

The blocker (S425): Word COM `doc.Paragraphs(pi)` order does NOT map cleanly
to literal XML `<w:p>` order (tables, textbox stories, AlternateContent,
floating anchors perturb it), and the counts don't reconcile (Word 2515 vs
XML main-story 2343 for 3a4f). So index-based mapping fails.

This aligns the three views by TEXT (non-empty paragraphs are the anchors),
which sidesteps the count mismatch:
  - Word side: pipeline_data/pagination_word/<doc>.json (per-para i, text,
    page, y) — already COM-measured, no Word needed here.
  - XML side: walk word/document.xml main story (body order; recurse into
    w:tbl/w:tr/w:tc and w:sdt; SKIP w:txbxContent and mc:Fallback), record
    per-<w:p> text + structural flags (has_drawing, has_floating_anchor,
    n_br, in_table).
  - Oxi side (optional): the gdi --dump-layout per-para records.

Greedy LCS-ish text alignment on normalized non-empty paragraph text gives
anchor pairs; the EMPTY paragraphs between consecutive anchors are bracketed
so we can count empties per side in each gap (the real drift signal) and
inspect the XML structure of any gap.

Usage:
  python tools/metrics/para_align.py <doc_id> [--gap-near-text "個人情報保護"]
"""
from __future__ import annotations
import os, sys, json, glob, zipfile
import xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

REPO = r'c:\Users\ryuji\oxi-main'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
DOCS = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx')
WORD_DIR = os.path.join(REPO, 'pipeline_data', 'pagination_word')


def norm(s):
    return ''.join((s or '').split())


def docx_for(doc_id):
    for p in glob.glob(os.path.join(DOCS, '*.docx')):
        if os.path.basename(p).startswith(doc_id):
            return p
    return None


def xml_main_story(docx):
    with zipfile.ZipFile(docx) as z:
        root = ET.fromstring(z.open('word/document.xml').read())
    body = root.find(f'{{{W}}}body')
    SKIP = {f'{{{W}}}txbxContent', f'{{{MC}}}Fallback'}
    out = []

    def text_of(p):
        return ''.join(t.text or '' for t in p.iter(f'{{{W}}}t'))

    def walk(el, in_table):
        for ch in el:
            tag = ch.tag
            if tag in SKIP:
                continue
            if tag == f'{{{W}}}p':
                t = text_of(ch)
                out.append({
                    'xml_idx': len(out),
                    'text': t[:40],
                    'empty': not t.strip(),
                    'n_br': len(list(ch.iter(f'{{{W}}}br'))),
                    'has_drawing': len(list(ch.iter(f'{{{W}}}drawing'))) > 0,
                    'has_pict': len(list(ch.iter(f'{{{W}}}pict'))) > 0,
                    'has_anchor': len(list(ch.iter(f'{{{W}}}anchor'))) > 0,  # floating
                    'in_table': in_table,
                })
            elif tag == f'{{{W}}}tbl':
                walk(ch, True)
            else:
                walk(ch, in_table)
    walk(body, False)
    return out


def align(word_paras, xml_paras):
    """Greedy align by normalized non-empty text. Returns anchor pairs
    [(word_idx, xml_idx)] in order."""
    wi = [(i, norm(w['text'])) for i, w in enumerate(word_paras) if norm(w['text'])]
    xi = [(j, norm(x['text'])) for j, x in enumerate(xml_paras) if norm(x['text'])]
    pairs = []
    a = b = 0
    while a < len(wi) and b < len(xi):
        wt = wi[a][1]; xt = xi[b][1]
        if wt[:8] == xt[:8] or wt.startswith(xt[:6]) or xt.startswith(wt[:6]):
            pairs.append((wi[a][0], xi[b][0]))
            a += 1; b += 1
        else:
            # lookahead: try to resync within a small window
            found = False
            for d in range(1, 6):
                if b + d < len(xi) and (wt[:8] == xi[b+d][1][:8]):
                    b += d; found = True; break
                if a + d < len(wi) and (wi[a+d][1][:8] == xt[:8]):
                    a += d; found = True; break
            if not found:
                a += 1; b += 1
    return pairs


def main():
    doc_id = [x for x in sys.argv[1:] if not x.startswith('--')][0]
    gap_text = None
    for x in sys.argv[1:]:
        if x.startswith('--gap-near-text='):
            gap_text = x.split('=', 1)[1]
    wj = json.load(open(os.path.join(WORD_DIR, f'{doc_id}.json'), encoding='utf-8'))
    word_paras = wj['paragraphs']
    xml_paras = xml_main_story(docx_for(doc_id))
    print(f'Word paras: {len(word_paras)} (empty {sum(1 for w in word_paras if not norm(w["text"]))})')
    print(f'XML main-story paras: {len(xml_paras)} (empty {sum(1 for x in xml_paras if x["empty"])})')
    pairs = align(word_paras, xml_paras)
    print(f'aligned anchors (non-empty text matches): {len(pairs)}')

    # For each consecutive anchor pair, count empties in the Word gap vs XML gap.
    big = []
    for k in range(len(pairs) - 1):
        w0, x0 = pairs[k]; w1, x1 = pairs[k + 1]
        w_gap = w1 - w0 - 1   # paras strictly between (Word)
        x_gap = x1 - x0 - 1   # paras strictly between (XML)
        if w_gap != x_gap or w_gap >= 2:
            xfl = xml_paras[x0+1:x1]
            big.append({
                'word_i': word_paras[w0].get('i'), 'word_text': word_paras[w0]['text'][:18],
                'next_word_text': word_paras[w1]['text'][:18],
                'w_gap': w_gap, 'x_gap': x_gap,
                'xml_gap_flags': {
                    'br': sum(p['n_br'] for p in xfl),
                    'drawing': sum(p['has_drawing'] for p in xfl),
                    'pict': sum(p['has_pict'] for p in xfl),
                    'anchor': sum(p['has_anchor'] for p in xfl),
                    'in_table': sum(p['in_table'] for p in xfl),
                },
            })
    # Mismatched gaps (Word has more empties than XML) — the drift signal
    mm = [b for b in big if b['w_gap'] > b['x_gap']]
    print(f'\ngaps where Word has MORE paras than XML (Word empties not in XML): {len(mm)}')
    for b in mm[:12]:
        print(f"  word_i={b['word_i']} w_gap={b['w_gap']} x_gap={b['x_gap']} flags={b['xml_gap_flags']} "
              f"after={b['word_text']!r}")
    if gap_text:
        for b in big:
            if gap_text in b['next_word_text'] or gap_text in b['word_text']:
                print(f"\ngap near {gap_text!r}: {b}")


if __name__ == '__main__':
    main()
