# -*- coding: utf-8 -*-
"""S492 big-job: reliable per-line render-truth dy for 1ec1 blk3 textbox (Word EMF baseline vs
Oxi rendered glyph), matching by TEXT (UTF-8 file, NOT a bash heredoc -- cp932 rule). Confirms
Oxi textbox content is ~6pt too HIGH and whether it's uniform (offset) or per-line."""
import json
SC = 0.12
W = json.load(open(r'c:\tmp\1ec1_emf_pos.json', encoding='utf-8'))['records']
O = json.load(open(r'c:\tmp\_1ec1.json', encoding='utf-8'))['pages'][0]['elements']

# Word EMF records in the blk3 box region (baseline y 165-300pt, x 40-445pt)
wbox = []
for r in W:
    if not r['text'].strip():
        continue
    xpt, ypt = r['x'] * SC, r['y'] * SC
    if 40 <= xpt <= 445 and 165 <= ypt <= 305:
        wbox.append({'x': round(xpt, 1), 'base': round(ypt, 1), 'text': r['text']})
wbox.sort(key=lambda r: (r['base'], r['x']))

# Oxi text elements in box: glyph_top = y + text_y_off; baseline = top + 0.8*fs
obox = []
for e in O:
    if e.get('type') != 'text' or not e.get('text', '').strip():
        continue
    gtop = e['y'] + e.get('text_y_off', 0)
    if 40 <= e['x'] <= 445 and 168 <= gtop <= 315:
        obox.append({'x': round(e['x'], 1), 'top': round(gtop, 1), 'fs': e['font_size'],
                     'base': round(gtop + 0.8 * e['font_size'], 1), 'text': e['text']})
obox.sort(key=lambda r: (r['base'], r['x']))

# match each Oxi run to a Word run by first-char + x proximity (text reliable now, UTF-8)
print("Word EMF box runs (x, baseline, text):")
for r in wbox[:14]:
    print("  x=%.1f base=%.1f %r" % (r['x'], r['base'], r['text'][:12]))
print("\nper-run match (Oxi -> Word by first char + x within 8pt):")
dys = []
for o in obox:
    cands = [w for w in wbox if w['text'] and o['text'] and w['text'][0] == o['text'][0] and abs(w['x'] - o['x']) < 8]
    if cands:
        w = min(cands, key=lambda w: abs(w['base'] - o['base']))
        dy = round(w['base'] - o['base'], 1)
        dys.append(dy)
        print("  Oxi x=%.1f base=%.1f fs=%.0f %r -> Word base=%.1f dy(W-O)=%+.1f" % (
            o['x'], o['base'], o['fs'], o['text'][:8], w['base'], dy))
if dys:
    import statistics
    print("\ndy: n=%d mean=%+.2f median=%+.2f min=%+.1f max=%+.1f" % (
        len(dys), statistics.mean(dys), statistics.median(dys), min(dys), max(dys)))
    print("=> uniform positive = Oxi content uniformly too HIGH by that amount (push down)")
