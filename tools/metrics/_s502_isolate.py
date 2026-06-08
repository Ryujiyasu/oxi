# -*- coding: utf-8 -*-
"""S502 generic isolation (multi-page): match ON/OFF glyph dumps by index per page (S502
only shifts center/right grid-cell x), find changed runs, identify line, compare ON/OFF/Word
first-char x. Decides if S502 moves the cell toward (good) or away (bad) from Word.
Usage: python _s502_isolate.py <w.json> <on.json> <off.json> <out.txt>  cp932-safe."""
import json, io, sys

wj, onj, offj, out = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
W = json.load(io.open(wj, encoding='utf-8'))['pages']
ON = json.load(io.open(onj, encoding='utf-8'))['pages']
OFF = json.load(io.open(offj, encoding='utf-8'))['pages']
L = []
tot_on_err = tot_off_err = 0.0
nlines = 0
for pi in range(min(len(ON), len(OFF), len(W))):
    on = ON[pi]['glyphs']; off = OFF[pi]['glyphs']; wg = W[pi]['glyphs']
    if len(on) != len(off):
        L.append('p%d LEN MISMATCH on=%d off=%d' % (pi, len(on), len(off)))
        continue
    changed = [(i, on[i], off[i]) for i in range(len(on)) if abs(on[i]['x'] - off[i]['x']) > 0.2]
    if not changed:
        continue
    by_line = {}
    for i, o, f in changed:
        by_line.setdefault(round(f['baseline'], 0), []).append((i, o, f))
    wchars = [g['char'] for g in wg]
    for ly in sorted(by_line):
        grp = by_line[ly]
        txt = ''.join(o['char'] for _, o, _ in grp)
        xon0 = grp[0][1]['x']; xoff0 = grp[0][2]['x']
        wx0 = None
        for st in range(len(wchars) - len(txt) + 1):
            if ''.join(wchars[st:st + len(txt)]) == txt:
                wx0 = wg[st]['x']; break
        L.append('p%d line~%.0f n=%2d shift=%+.2f  text=%s' % (pi, ly, len(grp), xon0 - xoff0, txt[:24]))
        if wx0 is not None:
            eon = abs(xon0 - wx0); eoff = abs(xoff0 - wx0)
            tot_on_err += eon; tot_off_err += eoff; nlines += 1
            L.append('     WORD %.2f | ON %.2f (err %.2f) | OFF %.2f (err %.2f) -> %s closer' % (
                wx0, xon0, eon, xoff0, eoff, 'ON' if eon < eoff else 'OFF'))
        else:
            L.append('     (Word run not found)')
L.append('\n=== SUMMARY over %d matched changed lines ===' % nlines)
if nlines:
    L.append('mean |first-char err|: ON %.3f  OFF %.3f  -> S502 %s' % (
        tot_on_err / nlines, tot_off_err / nlines,
        'IMPROVES' if tot_on_err < tot_off_err else 'REGRESSES'))
with io.open(out, 'w', encoding='utf-8') as f:
    f.write('\n'.join(L) + '\n')
print('wrote', out)
