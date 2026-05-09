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


def main():
    docx = r'C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\bd90b00ab7a7_order_05.docx'
    pdf = r'C:\tmp\bd90b00ab7a7_order_05_word.pdf'

    print(f'Step 1: export PDF from Word')
    if not os.path.exists(pdf):
        export_pdf(docx, pdf)
    else:
        print(f'  PDF already exists: {pdf}')

    print(f'\nStep 2: measure Word COM paragraph anchors')
    anchors = measure_paragraph_anchors(docx)
    print(f'  Got {len(anchors)} anchors')

    print(f'\nStep 3: extract PDF glyphs via PyMuPDF')
    pages = extract_pdf_glyphs(pdf)
    print(f'  Got {len(pages)} pages, {sum(len(p["lines"]) for p in pages)} text lines')

    print(f'\nStep 4: match anchors to glyphs')
    matches = match_anchor_to_glyph(anchors, pages)
    print(f'  Matched {len(matches)} of {len(anchors)} paragraphs')

    print(f'\n=== bd90b00 — Word anchor vs PDF glyph y per paragraph ===')
    print(f'  {"i":>3} {"pg":>2} {"anchor_y":>8} {"glyph_y":>8} {"offset":>7} {"glyph_orig":>10} {"lh_rule":>7} {"lh_val":>6} {"fs":>5} text')
    for m in matches:
        marker = ''
        if abs(m['offset']) >= 1.0:
            marker = ' <<NON-ZERO>>'
        print(f'  {m["i"]:>3} {m["page"]:>2} {m["anchor_y"]:>8.2f} {m["glyph_top_y"]:>8.2f} {m["offset"]:>+7.2f} {m["glyph_origin_y"]:>10.2f} {m["lh_rule"]:>7} {m["lh_val"]:>6.1f} {m["fs"]:>5} {m["text"]!r}{marker}')


if __name__ == '__main__':
    main()
