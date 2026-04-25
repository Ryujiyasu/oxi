"""Scan all baseline DOCX files for Japanese punctuation adjacency patterns.

Count per-doc: how many paragraphs have adjacency compression opportunities?
(、。」）．，) followed by (、。「」（）．，) = Rule A
(「（) preceded by (「（) = Rule B

Output per-doc + summary table showing which docs are affected.
"""
import os
import glob
import zipfile
import re
import json

DOCX_DIR = "tools/golden-test/documents/docx"

CLOSING = set('、。」）．，')  # 、。」）．，
OPENING = set('「（')  # 「（
ANY_PUNCT = CLOSING | OPENING


def extract_text_runs(docx_path: str) -> str:
    """Extract concatenated text from all runs (rough — across all paragraphs)."""
    try:
        with zipfile.ZipFile(docx_path) as z:
            try:
                doc_xml = z.read('word/document.xml').decode('utf-8')
            except KeyError:
                return ''
    except zipfile.BadZipFile:
        return ''
    texts = re.findall(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', doc_xml)
    return '\n'.join(texts)


def scan_adjacencies(text: str) -> dict:
    """Count adjacency compression opportunities."""
    rule_a = 0  # closing followed by any punct
    rule_b = 0  # opening preceded by opening
    pairs = {}  # (prev, next) → count
    for i in range(len(text) - 1):
        a, b = text[i], text[i+1]
        if a in CLOSING and b in ANY_PUNCT:
            rule_a += 1
            key = f'{a}{b}'
            pairs[key] = pairs.get(key, 0) + 1
        if b in OPENING and a in OPENING:
            rule_b += 1
    return {'rule_a_count': rule_a, 'rule_b_count': rule_b, 'pairs': pairs}


def main():
    files = sorted(glob.glob(os.path.join(DOCX_DIR, '*.docx')))
    print(f"Scanning {len(files)} docs...")

    results = {}
    totals = {'rule_a': 0, 'rule_b': 0, 'pairs_total': {}}
    for f in files:
        name = os.path.splitext(os.path.basename(f))[0]
        text = extract_text_runs(f)
        res = scan_adjacencies(text)
        results[name] = res
        totals['rule_a'] += res['rule_a_count']
        totals['rule_b'] += res['rule_b_count']
        for k, v in res['pairs'].items():
            totals['pairs_total'][k] = totals['pairs_total'].get(k, 0) + v

    # Report: top docs by adjacency count
    sorted_docs = sorted(results.items(), key=lambda x: x[1]['rule_a_count'] + x[1]['rule_b_count'], reverse=True)
    print(f"\n=== Top 30 docs by adjacency count ===")
    print(f"{'Doc':<60s} {'RuleA':>6s} {'RuleB':>6s}")
    for (name, res) in sorted_docs[:30]:
        print(f"{name[:60]:<60s} {res['rule_a_count']:>6d} {res['rule_b_count']:>6d}")

    # Aggregate pair counts
    print(f"\n=== Most common adjacency pairs (baseline total) ===")
    sorted_pairs = sorted(totals['pairs_total'].items(), key=lambda x: x[1], reverse=True)
    for (pair, count) in sorted_pairs[:20]:
        print(f"  '{pair}': {count}")

    print(f"\n=== Grand totals ===")
    print(f"  Rule A (closing+any_punct): {totals['rule_a']}")
    print(f"  Rule B (opening+opening):   {totals['rule_b']}")
    print(f"  Docs with any adjacency:    {sum(1 for r in results.values() if r['rule_a_count']+r['rule_b_count'] > 0)}")
    print(f"  Docs with Rule A:           {sum(1 for r in results.values() if r['rule_a_count'] > 0)}")
    print(f"  Docs with Rule B:           {sum(1 for r in results.values() if r['rule_b_count'] > 0)}")

    # Save
    out = 'pipeline_data/baseline_adjacency_scan.json'
    with open(out, 'w', encoding='utf-8') as f:
        json.dump({'per_doc': results, 'totals': totals}, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {out}")


if __name__ == "__main__":
    main()
