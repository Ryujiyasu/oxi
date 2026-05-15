"""COM measurement: does Word apply space_before at the top of each page?

For a set of docs, for each page, find the first non-empty paragraph and
report (y, y - top_margin, sb, line_spacing, text). If y - top_margin ≈ sb,
Word is APPLYING sb at page top; if y - top_margin ≈ 0, Word is SUPPRESSING.

Hypothesis (d4d126): Word applies sb at top of page 2+ when document is in
Far East layout mode (`<w:useFELayout/>` in settings/compat). Oxi
unconditionally suppresses, causing cumulative drift.

Outputs per-doc JSON to pipeline_data/first_on_page_sb_<doc_id>.json and
a summary table to stdout.
"""
import json, sys, os
import win32com.client

DOCS = [
    'tools/golden-test/documents/docx/d4d126dfe1d9_tokumei_08_01-3.docx',
    'tools/golden-test/documents/docx/ed025cbecffb_index-23.docx',
    # 3 PASS docs to verify hypothesis doesn't regress them:
    'tools/golden-test/documents/docx/6514f214e482_tokumei_08_01-2.docx',
    'tools/golden-test/documents/docx/31420af1a08f_tokumei_08_07.docx',
    'tools/golden-test/documents/docx/15076df085f5_tokumei_08_09.docx',
]
TOP_MARGIN = 72.0  # pt (assume 1440tw = 72pt for all our test docs)

def measure(docx_path):
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc_id = os.path.basename(docx_path).split('_')[0]
    doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    rows = []
    try:
        n = doc.Paragraphs.Count
        # Find first non-empty paragraph on each page
        first_on_page = {}
        for wi in range(1, n + 1):
            p = doc.Paragraphs(wi)
            rng = p.Range
            text = (rng.Text or '').replace('\r','').replace('\x07','').strip()
            # Use collapsed-start range for accurate page/y
            start_rng = doc.Range(rng.Start, rng.Start)
            page = int(start_rng.Information(3))
            y    = float(start_rng.Information(6))
            if y > 800: continue  # bogus
            if page not in first_on_page:
                first_on_page[page] = dict(wi=wi, page=page, y=y,
                                          sb=float(p.SpaceBefore),
                                          ls=float(p.LineSpacing),
                                          lsr=int(p.LineSpacingRule),
                                          text=text[:50],
                                          y_off=y - TOP_MARGIN)
        rows = [first_on_page[p] for p in sorted(first_on_page.keys())]
        return rows
    finally:
        doc.Close(SaveChanges=False)

def main():
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.Quit()  # ensure clean state

    all_results = {}
    for docx in DOCS:
        if not os.path.exists(docx):
            print(f'MISSING: {docx}')
            continue
        doc_id = os.path.basename(docx).split('_')[0]
        print(f'\n=== {doc_id} ({os.path.basename(docx)}) ===')
        try:
            rows = measure(docx)
        except Exception as e:
            print(f'  ERROR: {e}')
            continue
        # Print summary
        print(f'{"page":>4} {"wi":>4} {"y":>7} {"y_off":>7} {"sb":>6} {"ls":>5} text')
        applied_count = 0
        suppressed_count = 0
        for r in rows:
            print(f'{r["page"]:>4} {r["wi"]:>4} {r["y"]:>7.2f} {r["y_off"]:>+7.2f} {r["sb"]:>6.2f} {r["ls"]:>5.1f} {r["text"]!r}')
            # Classify: if y_off > sb-0.5 (within tolerance of sb), Word applied sb
            if r['page'] == 1:
                continue
            if r['sb'] > 0.01:
                # Compare y_off to sb. Tolerance ~1pt.
                if r['y_off'] > r['sb'] - 1.0:
                    applied_count += 1
                else:
                    suppressed_count += 1
        verdict = 'APPLIED' if applied_count >= suppressed_count and applied_count > 0 else ('SUPPRESSED' if suppressed_count > 0 else 'INDETERMINATE')
        print(f'  → page-top sb verdict: {verdict} (applied={applied_count}, suppressed={suppressed_count})')
        all_results[doc_id] = dict(rows=rows, verdict=verdict)

    out = os.path.abspath('pipeline_data/first_on_page_sb_survey.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print(f'\nWrote {out}')

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    main()
