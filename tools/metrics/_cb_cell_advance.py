#!/usr/bin/env python3
"""Per-char fullwidth advance: Oxi-dump vs Word-PDF, char-stream aligned.

Measures whether Oxi's cell/body fullwidth-CJK advance matches Word's.
Oxi advance = element 'w'. Word advance = next_char.origin.x - this.origin.x
(excludes line-end ink-width artifact). Matches lines by char-stream difflib.

Usage: _cb_cell_advance.py <dump.json> <word.pdf> [--fs 10.5]
"""
import sys, json, io, statistics

def decode(s):
    try: return s.encode('latin1').decode('cp932')
    except Exception: return s

def is_full(c):
    o = ord(c)
    return o > 0x2000 and not (0xFF61 <= o <= 0xFF9F)  # exclude halfwidth katakana

YAK = set('、。，．・「」『』（）〔〕【】：；')

def oxi_adv(dump_path, fs):
    d = json.load(io.open(dump_path, encoding='utf-8'))
    advs = []
    for pg in d['pages']:
        for e in pg['elements']:
            if e.get('type') != 'text': continue
            if abs(e.get('font_size', 0) - fs) > 0.1: continue
            c = decode(e['text'])
            if len(c) != 1 or not is_full(c): continue
            advs.append((c, e['w']))
    return advs

def word_adv(pdf_path, fs):
    import fitz
    doc = fitz.open(pdf_path)
    advs = []
    for page in doc:
        d = page.get_text('rawdict')
        for block in d['blocks']:
            for line in block.get('lines', []):
                for span in line.get('spans', []):
                    if abs(span.get('size', 0) - fs) > 0.6: continue
                    chs = span.get('chars', [])
                    for i in range(len(chs) - 1):  # skip last (no successor)
                        c = chs[i]['c']
                        if not is_full(c): continue
                        adv = chs[i+1]['origin'][0] - chs[i]['origin'][0]
                        if 0 < adv < fs * 2:
                            advs.append((c, adv))
    return advs

def stats(advs, label):
    kanji = [a for c, a in advs if c not in YAK]
    yak = [a for c, a in advs if c in YAK]
    def s(v):
        if not v: return 'none'
        v2 = sorted(v)
        return f"n={len(v)} med={statistics.median(v):.3f} mean={statistics.mean(v):.3f} min={v2[0]:.3f} max={v2[-1]:.3f}"
    print(f"{label}: kanji/kana {s(kanji)}")
    print(f"{label}: yakumono  {s(yak)}")

if __name__ == '__main__':
    dump, pdf = sys.argv[1], sys.argv[2]
    fs = float(sys.argv[sys.argv.index('--fs')+1]) if '--fs' in sys.argv else 10.5
    print(f"--- fs={fs} ---")
    stats(oxi_adv(dump, fs), 'OXI ')
    stats(word_adv(pdf, fs), 'WORD')
