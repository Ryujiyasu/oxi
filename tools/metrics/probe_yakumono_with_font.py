"""COM-probe whether Word compresses yakumono pairs in ed025c, d4d126,
e3c545 (positive control), 0e7af1 (negative control).

R7.41 prerequisite: R7.38 memo claims ed025c (-0.37) and d4d126 (-0.21)
regressed under `cjk_family.contains("Meiryo")` gate, yet XML shows zero
Meiryo references in either doc and theme-resolves to 游明朝 (Yu Mincho).
Need ground truth: does Word actually compress yakumono pairs in these
docs?

Per-measurement we also capture Range.Font.NameFarEast so we can correlate
compression with the actually-resolved East-Asian font.

Output: pipeline_data/probe_yakumono_with_font.json
"""
import json
import os
import sys
import time

import win32com.client
import pythoncom

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def com_retry(fn, *args, max_attempts=8, base_delay=0.2, **kwargs):
    """Retry COM call on RPC_E_CALL_REJECTED with exponential backoff."""
    delay = base_delay
    for attempt in range(max_attempts):
        try:
            return fn(*args, **kwargs)
        except pythoncom.com_error as e:  # noqa: PERF203
            if attempt == max_attempts - 1:
                raise
            time.sleep(delay)
            delay = min(delay * 2.0, 2.5)
    return None


DOCS = {
    'ed025c':  'tools/golden-test/documents/docx/ed025cbecffb_index-23.docx',
    'd4d126':  'tools/golden-test/documents/docx/d4d126dfe1d9_tokumei_08_01-3.docx',
    'e3c545':  'tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx',
    '0e7af1':  'tools/golden-test/documents/docx/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx',
    'b837808': 'tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx',
}

# Closing-class yakumono chars (matches Oxi's is_yakumono_closing intent).
CLOSING = set('、。」』）〕】》〉｝］．，')
OPENING = set('「『（〔【《〈｛［')

# Cap per doc to keep runtime sane: at most this many yakumono pairs per font.
PAIRS_PER_FONT_CAP = 60
# Also cap paragraph scan depth (early-exit when caps hit for all common fonts).
MAX_PARAS_SCANNED = 600


def measure_doc(word, doc_id, docx_path):
    abs_path = os.path.abspath(docx_path)
    print(f'  Opening {abs_path} ...', flush=True)
    doc = com_retry(word.Documents.Open, abs_path, ReadOnly=True)
    results = []
    font_counts = {}
    try:
        time.sleep(0.5)  # let Word settle after Open
        n_paras = com_retry(lambda: doc.Paragraphs.Count)
        scan_to = min(n_paras, MAX_PARAS_SCANNED)
        print(f'  {n_paras} paragraphs (scanning first {scan_to})', flush=True)
        for p_idx in range(1, scan_to + 1):
            if p_idx % 100 == 0:
                print(f'    para {p_idx}/{scan_to} — collected {len(results)} measurements', flush=True)
            try:
                para = com_retry(lambda: doc.Paragraphs(p_idx))
            except Exception:
                continue
            try:
                r = com_retry(lambda: para.Range)
                r_start = int(com_retry(lambda: r.Start))
                text = com_retry(lambda: r.Text) or ''
            except Exception:
                continue
            n = len(text)
            if n < 2:
                continue
            # Find candidate yakumono pair positions
            pair_positions = []
            for i in range(n - 1):
                if text[i] in CLOSING or text[i] in OPENING:
                    pair_positions.append(i)
            if not pair_positions:
                continue
            # Measure
            for i in pair_positions:
                if 0 <= i and i + 1 < n:
                    try:
                        sub_i = com_retry(lambda: doc.Range(r_start + i, r_start + i + 1))
                        sub_i1 = com_retry(lambda: doc.Range(r_start + i + 1, r_start + i + 2))
                        x0 = com_retry(lambda: sub_i.Information(5))
                        y0 = com_retry(lambda: sub_i.Information(6))
                        x1 = com_retry(lambda: sub_i1.Information(5))
                        y1 = com_retry(lambda: sub_i1.Information(6))
                        f = com_retry(lambda: sub_i.Font)
                        fontEA = com_retry(lambda: f.NameFarEast) or ''
                        fontAsc = com_retry(lambda: f.NameAscii) or ''
                        fontSize = float(com_retry(lambda: f.Size) or 0)
                    except Exception:
                        continue
                    if abs(y0 - y1) > 3:
                        # Different lines — skip
                        continue
                    width_prev = x1 - x0
                    pair_str = text[i] + text[i + 1]
                    key = (fontEA, round(fontSize, 1))
                    font_counts[key] = font_counts.get(key, 0) + 1
                    if font_counts[key] > PAIRS_PER_FONT_CAP:
                        continue
                    results.append({
                        'para_idx': p_idx,
                        'char_idx': i,
                        'pair': pair_str,
                        'width_prev_pt': round(width_prev, 2),
                        'y': round(y0, 1),
                        'fontEA': fontEA,
                        'fontAscii': fontAsc,
                        'fontSize': fontSize,
                    })
    finally:
        try:
            com_retry(doc.Close, False)
        except Exception as e:
            print(f'  WARN: Close failed: {e}', flush=True)
    return results


def classify(width, font_size, tol=0.6):
    """Classify width relative to fullwidth=font_size, half=font_size/2."""
    if font_size <= 0:
        return f'NOSIZE({width:.2f})'
    full = font_size
    half = font_size * 0.5
    two_thirds = font_size * 2.0 / 3.0
    if abs(width - half) < tol:
        return 'HALF'
    if abs(width - two_thirds) < tol:
        return 'TWO_THIRDS'
    if abs(width - full) < tol:
        return 'FULL'
    return f'OTHER({width:.2f})'


def main():
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    out = {}
    try:
        for doc_id, path in DOCS.items():
            print(f'\n=== {doc_id} ===', flush=True)
            measurements = measure_doc(word, doc_id, path)
            print(f'  -> {len(measurements)} measurements', flush=True)
            out[doc_id] = measurements
    finally:
        word.Quit()
    os.makedirs('pipeline_data', exist_ok=True)
    out_path = 'pipeline_data/probe_yakumono_with_font.json'
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(out, f, indent=2, ensure_ascii=False)
    print(f'\nSaved {out_path}', flush=True)

    # Summary: per (doc, font, pair-class) -> compression rate
    print('\n== Compression rate by (doc, fontEA, fontSize, pair-class) ==')
    print(f'{"doc":<10s} {"fontEA":<16s} {"size":>5s} {"pair":<6s} '
          f'{"HALF":>5s} {"2/3":>5s} {"FULL":>5s} {"OTHER":>6s} {"total":>6s}')
    for doc_id, measurements in out.items():
        groups = {}
        for m in measurements:
            cls = classify(m['width_prev_pt'], m['fontSize'])
            cls_bucket = 'HALF' if cls == 'HALF' else 'TWO_THIRDS' if cls == 'TWO_THIRDS' else 'FULL' if cls == 'FULL' else 'OTHER'
            pair_class = ('closing' if m['pair'][0] in CLOSING
                          else 'opening' if m['pair'][0] in OPENING
                          else 'other')
            key = (m['fontEA'][:14], round(m['fontSize'], 1), pair_class)
            d = groups.setdefault(key, {'HALF': 0, 'TWO_THIRDS': 0, 'FULL': 0, 'OTHER': 0})
            d[cls_bucket] += 1
        for key, d in sorted(groups.items()):
            total = sum(d.values())
            font_ea, size, pair_class = key
            print(f'{doc_id:<10s} {font_ea:<16s} {size:>5.1f} {pair_class:<6s} '
                  f'{d["HALF"]:>5d} {d["TWO_THIRDS"]:>5d} {d["FULL"]:>5d} {d["OTHER"]:>6d} {total:>6d}')


if __name__ == '__main__':
    sys.exit(main())
