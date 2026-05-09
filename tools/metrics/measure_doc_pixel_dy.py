"""Day 32 part 16 — Generalized pixel-level dy measurement across all
paragraphs of a document.

Builds on Day 32 part 15 (`measure_oxi_png_glyph.py`) which proved
pixel-level Oxi-Word match for 2 hardcoded paragraphs in bd90b00.

This tool scans:
- Oxi PNG (per page): pixel glyph_top per paragraph, scanning narrow
  x-range of the first character from layout JSON
- Word PDF: glyph bbox_y0 per line, matched to paragraph via
  text-prefix on same page

Outputs JSON: per-paragraph (page, oxi_pixel_glyph_top_pt,
word_pdf_glyph_top_pt, dy_pt, fs, lh, text_prefix). Also computes
cumulative drift trajectory.

Why: Class A docs (bd90b00, de6e, db9ca, d77a58) fail Phase 1.
Day 32 part 1 reported +9.12pt cumulative drift on bd90b00 備考
(pi=37) using Information(6) — but that was anchor-based and
methodology-flawed. We need pixel-level cumulative drift to know
whether Class A failure is positional drift (need position
compensation) or pure threshold mismatch (need page-break formula
fix).
"""
from __future__ import annotations
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')


def find_glyph_top(img_arr, x_range, y_range, threshold=180, min_dark_pixels=3):
    """Topmost row in y_range with >= min_dark_pixels < threshold in x_range."""
    import numpy as np
    x_start, x_end = x_range
    y_start, y_end = y_range
    if x_end <= x_start or y_end <= y_start:
        return None
    for y in range(y_start, y_end):
        if y < 0 or y >= img_arr.shape[0]: continue
        row = img_arr[y, x_start:x_end]
        n_dark = int(np.sum(row < threshold))
        if n_dark >= min_dark_pixels:
            return y
    return None


def collect_paragraph_starts(layout):
    """For each (page_idx, para_idx), gather first-line elements and concatenate text.

    Each character is a separate text element in layout JSON. To get the full
    first line of a paragraph, group by (page, pi, y) and concatenate sorted
    by x.
    """
    # First pass: index all text elements by (page, pi)
    by_para = {}  # (page_idx, pi) -> list of elements
    for page_idx, page in enumerate(layout['pages']):
        for el in page.get('elements', []):
            if el.get('type') != 'text': continue
            pi = el.get('para_idx')
            if pi is None: continue
            text = el.get('text', '')
            if text == '' or text is None: continue
            by_para.setdefault((page_idx, pi), []).append(el)

    out = {}
    for (page_idx, pi), elems in by_para.items():
        # Find min y → first line
        ys = [e['y'] for e in elems]
        min_y = min(ys)
        # Group by y within tolerance (different lines may differ by ~lh)
        first_line = [e for e in elems if abs(e['y'] - min_y) < 1.0]
        first_line.sort(key=lambda e: e['x'])
        full_text = ''.join(e.get('text', '') for e in first_line)
        # Strip leading whitespace for matching
        text_for_match = full_text.lstrip()
        if not text_for_match: continue
        first = first_line[0]
        out[(page_idx, pi)] = {
            'page': page_idx + 1,
            'pi': pi,
            'y': min_y,
            'x': first['x'],
            'text': full_text,
            'text_match': text_for_match,
            'fs': first.get('font_size', 0) or first.get('fs', 0),
        }
    return out


def extract_pdf_glyphs(pdf_path):
    """Per-page list of {bbox_y0, text, span_size}."""
    import fitz
    doc = fitz.open(pdf_path)
    pages = []
    for page_idx in range(doc.page_count):
        page = doc.load_page(page_idx)
        d = page.get_text('dict')
        lines = []
        for block in d['blocks']:
            if block['type'] != 0: continue
            for line in block.get('lines', []):
                bbox = line['bbox']
                spans = line.get('spans', [])
                if not spans: continue
                first = spans[0]
                text = ''.join(s['text'] for s in spans)
                lines.append({
                    'bbox_y0': round(bbox[1], 3),
                    'bbox_y1': round(bbox[3], 3),
                    'text': text,
                    'fs': first.get('size', 0),
                })
        pages.append(lines)
    doc.close()
    return pages


def match_pdf_line(oxi_para, pdf_lines):
    """Match oxi paragraph first line to PDF line by text prefix + y proximity.

    Use longest possible prefix to discriminate; tiebreak by y proximity.
    """
    target = oxi_para.get('text_match', oxi_para.get('text', ''))
    target = target.replace(' ', '').replace('　', '')  # strip whitespace
    if not target: return None
    # Try with descending prefix length
    for prefix_len in (12, 8, 5, 3, 2):
        if prefix_len > len(target): continue
        prefix = target[:prefix_len]
        candidates = []
        for line in pdf_lines:
            line_text = line['text'].replace(' ', '').replace('　', '')
            if prefix in line_text[:prefix_len + 8]:
                candidates.append((abs(line['bbox_y0'] - oxi_para['y']), line))
        if candidates:
            candidates.sort(key=lambda x: x[0])
            # Reject if best y diff is huge (>30pt) — likely wrong match
            if candidates[0][0] > 30:
                continue
            return candidates[0][1]
    return None


def measure_doc(doc_id, docx_label, png_dir, layout_path, pdf_path):
    import numpy as np
    from PIL import Image

    print(f'\n=== {doc_id} ===')
    print(f'  layout: {layout_path}')
    print(f'  pdf:    {pdf_path}')

    # Load layout
    with open(layout_path, encoding='utf-8') as f:
        layout = json.load(f)
    para_starts = collect_paragraph_starts(layout)
    print(f'  layout: {len(layout["pages"])} pages, {len(para_starts)} para-page entries')

    # Load PDF
    pdf_pages = extract_pdf_glyphs(pdf_path)
    print(f'  pdf: {len(pdf_pages)} pages')

    # Per-page PNG load
    results = []
    for page_idx in range(len(layout['pages'])):
        page_num = page_idx + 1
        png_path = os.path.join(png_dir, f'{docx_label}_p{page_num}.png')
        if not os.path.exists(png_path):
            print(f'  page {page_num}: PNG missing ({png_path})')
            continue
        img = Image.open(png_path).convert('L')
        arr = np.array(img)
        # 1240 px / (8.27 in * 72 pt/in) ≈ 2.0833 px/pt
        scale = img.size[0] / (8.27 * 72.0)
        pdf_lines = pdf_pages[page_idx] if page_idx < len(pdf_pages) else []

        # Iterate paragraphs on this page (key (p_idx, pi), 0-indexed)
        page_paras = [(pi, info) for (p_idx, pi), info in para_starts.items() if p_idx == page_idx]
        page_paras.sort(key=lambda x: x[0])  # by pi

        for pi, info in page_paras:
            oxi_y = info['y']
            oxi_x = info['x']
            fs = info['fs'] or 10.5

            # Pixel scan
            y_start = int((oxi_y - 1.5) * scale)
            y_end = int((oxi_y + max(fs * 1.2, 12)) * scale)
            x_start = max(0, int(oxi_x * scale))
            x_end = min(arr.shape[1], int((oxi_x + max(fs, 10)) * scale))
            glyph_top_px = find_glyph_top(arr, (x_start, x_end), (y_start, y_end))
            oxi_pixel_glyph_y = glyph_top_px / scale if glyph_top_px is not None else None

            # PDF match
            pdf_line = match_pdf_line(info, pdf_lines)
            word_pdf_y = pdf_line['bbox_y0'] if pdf_line else None

            dy = None
            if oxi_pixel_glyph_y is not None and word_pdf_y is not None:
                dy = round(oxi_pixel_glyph_y - word_pdf_y, 3)

            results.append({
                'page': page_num,
                'pi': pi,
                'oxi_layout_y': round(oxi_y, 3),
                'oxi_pixel_glyph_y': round(oxi_pixel_glyph_y, 3) if oxi_pixel_glyph_y is not None else None,
                'word_pdf_glyph_y': word_pdf_y,
                'dy': dy,
                'fs': round(fs, 2),
                'text': info['text'][:20],
                'pdf_text': pdf_line['text'][:30] if pdf_line else None,
            })

    return results


def main():
    if len(sys.argv) < 2:
        print('Usage: measure_doc_pixel_dy.py <doc_label> [pdf_path] [layout_path]')
        print('  example: measure_doc_pixel_dy.py bd90b00ab7a7_order_05')
        sys.exit(1)

    docx_label = sys.argv[1]
    pdf_path = sys.argv[2] if len(sys.argv) > 2 else f'C:/tmp/{docx_label}_word.pdf'
    layout_path = sys.argv[3] if len(sys.argv) > 3 else f'C:/tmp/{docx_label}_v2_layout.json'

    doc_id = docx_label.split('_', 1)[0][:12]
    png_dir = 'pipeline_data/oxi_gdi_tmp'

    results = measure_doc(doc_id, docx_label, png_dir, layout_path, pdf_path)

    # Output summary
    print(f'\n{"page":>4} {"pi":>3} {"oxi_layout_y":>13} {"oxi_pix_gy":>11} {"word_pdf_y":>11} {"dy":>7} {"fs":>5} {"text":<22}')
    drifts = []
    for r in results:
        if r['dy'] is None:
            continue
        drifts.append(r['dy'])
        print(f'  {r["page"]:>3} {r["pi"]:>3} {r["oxi_layout_y"]:>13.2f} {r["oxi_pixel_glyph_y"]:>11.2f} {r["word_pdf_glyph_y"]:>11.2f} {r["dy"]:>+7.2f} {r["fs"]:>5.1f} {r["text"]:<22}')
    print(f'\n  matched: {len(drifts)} paragraphs with dy')
    if drifts:
        n = len(drifts)
        mn = sum(drifts) / n
        print(f'  dy: min={min(drifts):+.2f} max={max(drifts):+.2f} mean={mn:+.2f}')
        # Trajectory: cumulative drift indicator
        first_5 = drifts[:5]
        last_5 = drifts[-5:]
        if first_5 and last_5:
            print(f'  first 5 dy mean: {sum(first_5)/len(first_5):+.2f}')
            print(f'  last  5 dy mean: {sum(last_5)/len(last_5):+.2f}')

    # Save JSON
    out_path = f'pipeline_data/pixel_dy_{doc_id}.json'
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump({'doc_id': doc_id, 'docx_label': docx_label, 'results': results},
                  f, ensure_ascii=False, indent=2)
    print(f'\n  saved: {out_path}')


if __name__ == '__main__':
    main()
