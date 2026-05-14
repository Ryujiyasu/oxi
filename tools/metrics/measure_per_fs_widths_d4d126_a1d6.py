"""COM-measure Word's per-char advance for each fs in d4d126/a1d6 docGrid.

Goal (Day 37 autonomous loop): determine Word's per-char width formula as a
function of fs in linesAndChars docGrid (linePitch=292, charSpace=1453,
default_fs=10.5).

Output: pipeline_data/per_fs_widths_d4d126_a1d6.json

Method: iterate Word paragraphs (NOT XML), filter by font size, pick first
paragraph at each target fs that has a 1-line segment with >=10 CJK fullwidth
chars. Measure per-char x positions. Output stats.
"""
from __future__ import annotations
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOCS = [
    ('d4d126', 'tools/golden-test/documents/docx/d4d126dfe1d9_tokumei_08_01-3.docx'),
    ('a1d6', 'tools/golden-test/documents/docx/a1d6e4efa2e7_tokumei_08_01-4.docx'),
]
OUT = os.path.abspath('pipeline_data/per_fs_widths_d4d126_a1d6.json')

TARGET_FS = [8, 9, 9.5, 10, 11, 12]


def is_cjk_fullwidth(ch: str) -> bool:
    cp = ord(ch)
    if 0x4E00 <= cp <= 0x9FFF: return True  # Kanji
    if 0x3041 <= cp <= 0x309F: return True  # Hiragana
    if 0x30A1 <= cp <= 0x30FF: return True  # Katakana
    if 0x3000 <= cp <= 0x303F: return True  # CJK punctuation+space
    if 0xFF01 <= cp <= 0xFF60: return True  # Fullwidth Latin
    return False


def measure_para_widths(d, para):
    """Measure per-char positions for a single paragraph."""
    rng = para.Range
    full_text = (rng.Text or '').rstrip('\r\n\x07')
    if len(full_text) < 5:
        return None
    start = rng.Start
    chars = []
    for i, ch in enumerate(full_text):
        ch_rng = d.Range(start + i, start + i)
        try:
            x = ch_rng.Information(5)
            y = ch_rng.Information(6)
        except Exception:
            return None
        chars.append({'i': i, 'ch': ch, 'x': round(x, 3), 'y': round(y, 3)})
    from collections import defaultdict
    by_line = defaultdict(list)
    for c in chars:
        by_line[c['y']].append(c)
    lines = []
    for y in sorted(by_line.keys()):
        line_chars = by_line[y]
        if len(line_chars) < 2:
            continue
        n_chars = len(line_chars)
        line_width = line_chars[-1]['x'] - line_chars[0]['x']
        avg_cw = line_width / (n_chars - 1) if n_chars > 1 else None
        deltas = []
        for j in range(n_chars - 1):
            deltas.append(round(line_chars[j+1]['x'] - line_chars[j]['x'], 3))
        delta_counter = {}
        for d_val in deltas:
            delta_counter[d_val] = delta_counter.get(d_val, 0) + 1
        # Also: count of fullwidth chars in line (filter out Latin)
        n_fullwidth = sum(1 for c in line_chars if is_cjk_fullwidth(c['ch']))
        lines.append({
            'y': y,
            'n_chars': n_chars,
            'n_fullwidth': n_fullwidth,
            'text': ''.join(c['ch'] for c in line_chars)[:30],
            'line_width': round(line_width, 3),
            'avg_cw': round(avg_cw, 4) if avg_cw is not None else None,
            'delta_counts': dict(sorted(delta_counter.items())),
        })
    para_format = para.Format
    fnt = rng.Font
    return {
        'word_para_idx': para.Range.Information(10) if para else None,
        'text': full_text[:60],
        'font_size_pt': fnt.Size,
        'font_name': fnt.NameFarEast or fnt.Name,
        'spacing': fnt.Spacing,
        'left_indent_pt': para_format.LeftIndent,
        'first_line_indent_pt': para_format.FirstLineIndent,
        'lines': lines,
    }


def main():
    out = {'measurements': []}
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for doc_id, docx_path in DOCS:
            full_path = os.path.abspath(docx_path)
            print(f"\n=== {doc_id} ===")
            d = word.Documents.Open(full_path, ReadOnly=True)
            try:
                n_paras = d.Paragraphs.Count
                # Per-fs: take first paragraph at that fs with >=8 fullwidth chars in any line
                found_per_fs = {fs: None for fs in TARGET_FS}
                for pi in range(1, min(n_paras, 800) + 1):
                    para = d.Paragraphs(pi)
                    fs = para.Range.Font.Size
                    # Match to nearest target fs
                    target = None
                    for tfs in TARGET_FS:
                        if abs(fs - tfs) < 0.1:
                            target = tfs
                            break
                    if target is None or found_per_fs[target] is not None:
                        continue
                    txt = (para.Range.Text or '').rstrip('\r\n\x07')
                    if len(txt) < 10:
                        continue
                    n_fw = sum(1 for c in txt if is_cjk_fullwidth(c))
                    if n_fw < 8:
                        continue  # need enough CJK chars
                    # Measure
                    data = measure_para_widths(d, para)
                    if data is None or not data['lines']:
                        continue
                    # Find a line with >=8 fullwidth chars
                    good_line = next((l for l in data['lines'] if l['n_fullwidth'] >= 8), None)
                    if good_line is None:
                        continue
                    data['doc_id'] = doc_id
                    data['expected_fs'] = target
                    data['word_para_idx'] = pi
                    found_per_fs[target] = data
                    print(f"  fs={target} wi={pi} reported_fs={fs:.2f} line: n_fw={good_line['n_fullwidth']} avg_cw={good_line['avg_cw']} deltas={good_line['delta_counts']} text={good_line['text']!r}")
                    if all(v is not None for v in found_per_fs.values()):
                        break
                for fs, data in found_per_fs.items():
                    if data is not None:
                        out['measurements'].append(data)
            finally:
                d.Close(SaveChanges=False)
    finally:
        word.Quit()
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT}")


if __name__ == '__main__':
    main()
