"""COM-probe whether Word actually compresses yakumono pairs in 0e7af + 683ff.

Hypothesis to test: the V_CP/V_COMPAT15 8x8 fixtures showed Word compresses
ALL closing-yakumono pairs unconditionally. But applying that rule via Oxi
(`yakumono_enabled=true`) collapsed 0e7af pp.2-7 (-0.19 to -0.29 each) and
683ff pp.1-3 (-0.04 to -0.25). Either:
  (A) Word does NOT compress yakumono pairs in 0e7af/683ff (some
      context-discriminator excludes them) — so Oxi opening the gate
      created compression that Word does not apply.
  (B) Word DOES compress them, but Oxi's compression implementation has
      a bug that breaks layout when applied broadly to multi-pair lines.

This script measures per-character X positions of the yakumono pairs in
0e7af + 683ff, compares (next.x - prev.x) widths to the
"compressed=5.25pt" / "full=10.50pt" thresholds, and reports per-pair
classification. If (A) holds, we should see lots of full-width (10.50)
in the regressed docs vs always-5.25 in the V_CP fixtures.

Run: python tools/metrics/probe_0e7af_yakumono.py
Output: pipeline_data/probe_0e7af_yakumono.json
"""
import json
import os
import sys

import win32com.client


DOCS = [
    'tools/golden-test/documents/docx/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx',
    'tools/golden-test/documents/docx/683ffcab86e2_20230331_resources_open_data_contract_addon_00.docx',
]

# Closing-class yakumono chars (matching Oxi's is_yakumono_closing intent).
CLOSING = set('、。」）．，)')


def measure_doc(word, docx_path):
    abs_path = os.path.abspath(docx_path)
    doc = word.Documents.Open(abs_path, ReadOnly=True)
    try:
        # Walk all paragraphs, find yakumono pairs (closing followed by anything),
        # measure per-char widths around the pair.
        results = []
        for p_idx in range(1, doc.Paragraphs.Count + 1):
            try:
                para = doc.Paragraphs(p_idx)
            except Exception:
                continue
            r = para.Range
            text = r.Text
            n = len(text)
            if n < 2:
                continue
            # Find pair positions in this paragraph
            pair_positions = []
            for i in range(n - 1):
                if text[i] in CLOSING:
                    pair_positions.append(i)
            if not pair_positions:
                continue
            # Measure x for each char around pairs
            for i in pair_positions:
                # Need positions of chars i-1 (if exists), i, i+1, i+2 (if exists)
                positions = {}
                for off in [-1, 0, 1, 2]:
                    j = i + off
                    if j < 0 or j >= n:
                        continue
                    sub = doc.Range(r.Start + j, r.Start + j + 1)
                    try:
                        positions[off] = {
                            'ch': text[j],
                            'x': sub.Information(5),
                            'y': sub.Information(6),
                        }
                    except Exception:
                        positions[off] = {'ch': text[j], 'x': None, 'y': None}
                # Only compute width if i and i+1 are on the same line
                if 0 in positions and 1 in positions:
                    p0 = positions[0]
                    p1 = positions[1]
                    if p0['x'] is not None and p1['x'] is not None and p0['y'] is not None and p1['y'] is not None:
                        if abs(p0['y'] - p1['y']) < 3:
                            width_prev = p1['x'] - p0['x']
                            results.append({
                                'para_idx': p_idx,
                                'char_idx': i,
                                'pair': p0['ch'] + p1['ch'],
                                'prev_ch': p0['ch'],
                                'next_ch': p1['ch'],
                                'width_prev': round(width_prev, 2),  # width of prev_ch (the closing yakumono)
                                'y': round(p0['y'], 1),
                            })
        return results
    finally:
        doc.Close(False)


def classify(width, full=10.5, compressed=5.25, tol=0.5):
    if abs(width - compressed) < tol:
        return 'COMPRESSED'
    if abs(width - full) < tol:
        return 'FULL'
    return f'OTHER({width:.2f})'


def main():
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    out = {}
    try:
        for docx_path in DOCS:
            doc_id = os.path.splitext(os.path.basename(docx_path))[0]
            print(f'Measuring {doc_id} ...', flush=True)
            measurements = measure_doc(word, docx_path)
            print(f'  -> {len(measurements)} pair measurements', flush=True)
            out[doc_id] = measurements
    finally:
        word.Quit()
    os.makedirs('pipeline_data', exist_ok=True)
    out_path = 'pipeline_data/probe_0e7af_yakumono.json'
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(out, f, indent=2, ensure_ascii=False)
    print(f'Saved {out_path}', flush=True)

    # Summary classification
    print('\n== Per-pair classification by next-char class ==')
    print(f'{"doc":<60s} {"pair":<6s} {"COMPRESSED":>10s} {"FULL":>6s} {"OTHER":>6s}')
    for doc_id, measurements in out.items():
        # Group by pair string
        by_pair = {}
        for m in measurements:
            by_pair.setdefault(m['pair'], []).append(m)
        for pair, ms in sorted(by_pair.items(), key=lambda x: -len(x[1])):
            comp = sum(1 for m in ms if classify(m['width_prev']) == 'COMPRESSED')
            full = sum(1 for m in ms if classify(m['width_prev']) == 'FULL')
            other = len(ms) - comp - full
            print(f'{doc_id[:60]:<60s} {pair:<6s} {comp:>10d} {full:>6d} {other:>6d}')


if __name__ == '__main__':
    sys.exit(main())
