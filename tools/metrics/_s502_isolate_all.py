# -*- coding: utf-8 -*-
"""S502 position gate over the affected docs: for each, generate Word PDF glyphs + Oxi ON
+ Oxi OFF dumps, then isolate changed center lines and report mean first-char err ON vs OFF
(vs Word). The meaningful gate for this horizontal position fix (SSIM is position-blind at
this sub-page scale). cp932-safe: ASCII out file."""
import os, sys, json, io, subprocess, tempfile, glob

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
DOCS = ['29dc6e8943fe_order_01', '6514f214e482_tokumei_08_01-2', 'a1d6e4efa2e7_tokumei_08_01-4',
        'd4d126dfe1d9_tokumei_08_01-3', 'de6e32b5960b_tokumei_08_01-1']


def docx_for(did):
    for d in ['tools/golden-test/documents/docx', 'pipeline_data/docx']:
        for p in glob.glob(os.path.join(ROOT, d, did + '*.docx')):
            if os.path.splitext(os.path.basename(p))[0] == did:
                return p
        g = glob.glob(os.path.join(ROOT, d, did + '*.docx'))
        if g:
            return g[0]
    return None


def dump(dx, jp, disable):
    e = dict(os.environ)
    if disable:
        e['OXI_S502_DISABLE'] = '1'
    subprocess.run([DW, os.path.abspath(dx), tempfile.mktemp(dir='c:/tmp'), '150',
                    '--dump-glyphs=' + jp], capture_output=True, timeout=400, env=e)


def isolate(wj, onj, offj):
    W = json.load(io.open(wj, encoding='utf-8'))['pages']
    ON = json.load(io.open(onj, encoding='utf-8'))['pages']
    OFF = json.load(io.open(offj, encoding='utf-8'))['pages']
    eon = eoff = 0.0
    n = 0
    lines = []
    for pi in range(min(len(ON), len(OFF), len(W))):
        on = ON[pi]['glyphs']; off = OFF[pi]['glyphs']; wg = W[pi]['glyphs']
        if len(on) != len(off):
            continue
        changed = [(i, on[i], off[i]) for i in range(len(on)) if abs(on[i]['x'] - off[i]['x']) > 0.2]
        by = {}
        for i, o, f in changed:
            by.setdefault(round(f['baseline'], 0), []).append((i, o, f))
        wch = [g['char'] for g in wg]
        for ly in sorted(by):
            grp = by[ly]
            txt = ''.join(o['char'] for _, o, _ in grp)
            xon = grp[0][1]['x']; xoff = grp[0][2]['x']
            wx = None
            for st in range(len(wch) - len(txt) + 1):
                if ''.join(wch[st:st + len(txt)]) == txt:
                    wx = wg[st]['x']; break
            if wx is not None:
                eon += abs(xon - wx); eoff += abs(xoff - wx); n += 1
                lines.append('  p%d %-20s ON_err %.2f OFF_err %.2f' % (pi, txt[:20], abs(xon - wx), abs(xoff - wx)))
    return n, eon, eoff, lines


def main():
    out = ['S502 position gate (first-char err vs Word, ON=fix OFF=disabled)']
    tn = teon = teoff = 0
    for did in DOCS:
        dx = docx_for(did)
        if not dx:
            out.append('%s MISSING' % did); continue
        wj = 'c:/tmp/s502_%s_w.json' % did[:8]
        onj = 'c:/tmp/s502_%s_on.json' % did[:8]
        offj = 'c:/tmp/s502_%s_off.json' % did[:8]
        if not os.path.exists(wj):
            subprocess.run([sys.executable, os.path.join(ROOT, 'tools', 'metrics', 'word_pdf_glyphs.py'), dx, wj],
                           capture_output=True, timeout=400)
        dump(dx, onj, False); dump(dx, offj, True)
        n, eon, eoff, lines = isolate(wj, onj, offj)
        tn += n; teon += eon; teoff += eoff
        verdict = 'IMPROVES' if eon < eoff - 1e-6 else ('REGRESS' if eon > eoff + 1e-6 else 'flat')
        out.append('\n=== %s : %d changed lines, mean ON %.3f / OFF %.3f -> %s ===' % (
            did[:24], n, (eon / n if n else 0), (eoff / n if n else 0), verdict))
        out.extend(lines)
    out.append('\n==== TOTAL %d lines: mean ON %.3f / OFF %.3f -> %s ====' % (
        tn, (teon / tn if tn else 0), (teoff / tn if tn else 0),
        'IMPROVES' if teon < teoff else 'REGRESS'))
    with io.open('c:/tmp/_s502_posgate_out.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(out) + '\n')
    print('\n'.join(out))


if __name__ == '__main__':
    main()
