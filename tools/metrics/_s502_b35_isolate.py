# -*- coding: utf-8 -*-
"""S502 b35 isolation: ON and OFF glyph dumps have identical glyph sequences (S502 only
shifts center/right grid-cell x), so match by index, find the changed run, identify the
line (y), and compare ON/OFF/Word x for those glyphs. Decides whether Word centers b35's
negative-charSpace cell on the natural or grid-adjusted width. cp932-safe: ASCII out file."""
import json, io, statistics

ON = json.load(io.open('c:/tmp/b35_on.json', encoding='utf-8'))['pages'][0]['glyphs']
OFF = json.load(io.open('c:/tmp/b35_off.json', encoding='utf-8'))['pages'][0]['glyphs']
W = json.load(io.open('c:/tmp/b35_w.json', encoding='utf-8'))['pages'][0]['glyphs']

L = []
if len(ON) != len(OFF):
    L.append('LEN MISMATCH on=%d off=%d (S502 changed wrapping?!)' % (len(ON), len(OFF)))
else:
    L.append('glyph counts match: %d' % len(ON))
    # find changed glyphs (x differs)
    changed = [(i, ON[i], OFF[i]) for i in range(len(ON)) if abs(ON[i]['x'] - OFF[i]['x']) > 0.2]
    L.append('changed glyphs: %d' % len(changed))
    # group changed by their OFF y (baseline) -> line
    by_line = {}
    for i, o, f in changed:
        key = round(f['baseline'], 0)
        by_line.setdefault(key, []).append((i, o, f))
    for ly in sorted(by_line):
        grp = by_line[ly]
        txt = ''.join(o['char'] for _, o, _ in grp)
        xon0 = grp[0][1]['x']; xoff0 = grp[0][2]['x']
        L.append('\n== line baseline~%.0f  n=%d  shift_on-off=%+.2f ==' % (ly, len(grp), xon0 - xoff0))
        L.append('  text: %s' % txt)
        # find the same text run in Word glyphs near this y (Word baseline ~ off baseline + cal)
        # match by char content: find contiguous Word glyphs equal to txt
        wchars = [g['char'] for g in W]
        s = txt
        wx0 = None
        for st in range(len(wchars) - len(s) + 1):
            if ''.join(wchars[st:st + len(s)]) == s:
                wx0 = W[st]['x']; wy0 = W[st]['y']
                break
        if wx0 is not None:
            L.append('  first-char x:  WORD %.2f | ON %.2f | OFF %.2f' % (wx0, xon0, xoff0))
            L.append('  |ON-WORD| %.2f   |OFF-WORD| %.2f   -> %s closer' % (
                abs(xon0 - wx0), abs(xoff0 - wx0),
                'ON' if abs(xon0 - wx0) < abs(xoff0 - wx0) else 'OFF'))
        else:
            L.append('  (Word run not found by exact content match)')

with io.open('c:/tmp/_s502_b35_isolate_out.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(L) + '\n')
print('wrote c:/tmp/_s502_b35_isolate_out.txt')
