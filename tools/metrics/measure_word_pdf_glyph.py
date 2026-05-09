"""Day 32 part 13c — Word の実際の text glyph 描画 y を PDF 経由で測定.

COM Information(6) は paragraph anchor（line box top）を返すだけで、
line box 内の text glyph 位置は分からない。

代わりに Word でドキュメントを PDF に保存して、PDF 内の各 char の
(x, y) を fitz (PyMuPDF) で抽出する。これは Word が「実際に描画した」
位置なので、line box 内の text 配置を直接観測できる。

Hypothesis (Day 32 part 12): bd90b00 pi=11 (lh=Exact 16, fs=11.5):
- もし Word が text を line box top に置く → glyph_y = paragraph_anchor_y
- もし Word が text を line box bottom に置く → glyph_y = anchor + (lh-fs) = anchor + 4.5

bd90b00 を PDF に保存し、pi=11 の "年" 文字の y を抽出して比較。
"""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding='utf-8')


def export_pdf(docx_path, pdf_path):
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    try:
        # SaveAs2 with wdFormatPDF=17
        d.SaveAs2(os.path.abspath(pdf_path), FileFormat=17)
        print(f'  Exported PDF: {pdf_path}')
    finally:
        d.Close(False)
        word.Quit()


def measure_paragraph_anchors(docx_path):
    """Get Word COM paragraph anchor (Information(6)) for reference."""
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    anchors = []
    try:
        # bd90b00 has 82 paragraphs; sample first 50
        n = min(d.Paragraphs.Count, 50)
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            r = p.Range
            text = (r.Text or '').strip()
            cr_start = d.Range(r.Start, r.Start)
            try:
                pg = int(cr_start.Information(3))
                y = round(cr_start.Information(6), 2)
                lh_rule = p.Format.LineSpacingRule
                lh_val = round(p.Format.LineSpacing, 2)
                fs = r.Font.Size
            except Exception:
                pg, y, lh_rule, lh_val, fs = -1, -1, -1, -1, -1
            anchors.append({'i': i, 'page': pg, 'anchor_y': y,
                          'lh_rule': lh_rule, 'lh_val': lh_val, 'fs': fs,
                          'text': text[:40]})
    finally:
        d.Close(False)
        word.Quit()
    return anchors


def extract_pdf_glyphs(pdf_path):
    """Extract first glyph y for each line in each page using PyMuPDF."""
    import fitz
    doc = fitz.open(pdf_path)
    pages = []
    for page_idx in range(doc.page_count):
        page = doc.load_page(page_idx)
        # Get text dict with glyph-level positions
        d = page.get_text('dict')
        lines = []
        for block in d['blocks']:
            if block['type'] != 0:  # type 0 = text
                continue
            for line in block.get('lines', []):
                bbox = line['bbox']  # (x0, y0, x1, y1)
                # Get first span text
                spans = line.get('spans', [])
                if not spans:
                    continue
                first = spans[0]
                text = ''.join(s['text'] for s in spans)
                lines.append({
                    'page': page_idx + 1,
                    'bbox_y0': round(bbox[1], 2),  # top of glyph bbox
                    'bbox_y1': round(bbox[3], 2),  # bottom of glyph bbox
                    'span_origin_y': round(first.get('origin', [0,0])[1], 2),
                    'span_size': first.get('size'),
                    'text': text[:50],
                })
        pages.append({'page': page_idx + 1, 'page_height': page.rect.height, 'lines': lines})
    doc.close()
    return pages


def match_anchor_to_glyph(anchors, pages):
    """For each anchor (Word COM paragraph), find matching glyph line (by text prefix on same page)."""
    matches = []
    for a in anchors:
        if not a['text']:
            continue
        prefix = a['text'][:10]
        if a['page'] > len(pages):
            continue
        page = pages[a['page'] - 1]
        for line in page['lines']:
            if line['text'].startswith(prefix[:5]) or prefix[:5] in line['text'][:10]:
                glyph_y = line['bbox_y0']  # top of glyph bbox
                offset = round(glyph_y - a['anchor_y'], 2)
                matches.append({
                    'i': a['i'], 'page': a['page'],
                    'lh_rule': a['lh_rule'], 'lh_val': a['lh_val'], 'fs': a['fs'],
                    'anchor_y': a['anchor_y'],
                    'glyph_top_y': glyph_y,
                    'glyph_bottom_y': line['bbox_y1'],
                    'glyph_origin_y': line['span_origin_y'],
                    'offset': offset,
                    'text': a['text'][:40],
                })
                break
    return matches


def find_docx(doc_id, docx_dir):
    for f in os.listdir(docx_dir):
        if f.startswith(doc_id) and f.endswith('.docx'):
            return os.path.join(docx_dir, f)
    return None


def process(doc_id, docx_dir):
    docx = find_docx(doc_id, docx_dir)
    if not docx:
        print(f'  {doc_id}: NOT FOUND')
        return None
    label = os.path.splitext(os.path.basename(docx))[0]
    pdf = os.path.join(r'C:\tmp', f'{label}_word.pdf')
    if not os.path.exists(pdf):
        print(f'  exporting PDF...')
        export_pdf(docx, pdf)
    anchors = measure_paragraph_anchors(docx)
    pages = extract_pdf_glyphs(pdf)
    matches = match_anchor_to_glyph(anchors, pages)
    return {'doc_id': doc_id, 'matches': matches, 'n_anchors': len(anchors)}


def main():
    docx_dir = r'C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx'
    docs = ['bd90b00ab7a7', 'de6e32b5960b', 'db9ca18368cd', 'd77a58485f16']

    print('Re-measuring 4 Class A docs with PDF-glyph methodology...')
    all_results = {}
    for doc_id in docs:
        print(f'\n--- {doc_id} ---')
        r = process(doc_id, docx_dir)
        if not r:
            continue
        all_results[doc_id] = r
        ms = r['matches']
        print(f'  matched {len(ms)} of {r["n_anchors"]} sampled paragraphs')
        offsets = [m["offset"] for m in ms]
        if offsets:
            print(f'  offset range: {min(offsets):+.2f} to {max(offsets):+.2f} (mean {sum(offsets)/len(offsets):+.2f})')

    # Cross-doc comparison: how much offset is "natural Word behavior"?
    print('\n=== Cross-doc Word natural offset summary ===')
    print('  (offset = PDF glyph_top_y - Word anchor_y)')
    print('  if all offsets are similar across docs, that reflects Word internal line-box behavior')
    by_fs_lh = {}
    for doc_id, r in all_results.items():
        for m in r['matches']:
            key = (m['lh_rule'], m['lh_val'], m['fs'])
            by_fs_lh.setdefault(key, []).append(m['offset'])
    print(f'  {"lh_rule":>7} {"lh_val":>6} {"fs":>5} {"n":>3} {"mean":>8} {"min":>7} {"max":>7}')
    for key, offsets in sorted(by_fs_lh.items()):
        n = len(offsets)
        mn = sum(offsets)/n
        print(f'  {key[0]:>7} {key[1]:>6.1f} {key[2]:>5} {n:>3} {mn:>+8.2f} {min(offsets):>+7.2f} {max(offsets):>+7.2f}')


if __name__ == '__main__':
    main()
