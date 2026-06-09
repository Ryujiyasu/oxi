# -*- coding: utf-8 -*-
"""S516: scan a doc's Oxi dwrite glyph dump for WITHIN-LINE baseline splits = the mixed-font
renderer baseline bug (Yu Mincho runs sit ~3.3pt lower than Times/MS runs on the same line).
A 'line' = glyphs within a visual band; report bands that contain >1 distinct baseline (>1.5pt
apart) and the font of each sub-baseline. cp932-safe: UTF-8 file, results to file, ASCII out."""
import os, sys, json, subprocess, io, collections
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')

def font_tag(fam):
    cps = [ord(c) for c in fam]
    if any(c > 0x3000 for c in cps):
        if '游明朝' in fam: return 'YuMincho'
        if '游ゴシック' in fam: return 'YuGothic'
        if 'ＭＳ' in fam or 'MS' in fam: return 'MS-CJK'
        return 'CJK:' + '+'.join(hex(c) for c in cps[:3])
    return fam

def scan(docx):
    pre = os.path.join('c:/tmp', os.path.splitext(os.path.basename(docx))[0] + '_bs')
    gj = pre + '_glyphs.json'
    subprocess.run([EXE, os.path.abspath(docx), pre, '72', '--dump-glyphs=' + gj], capture_output=True, text=True)
    if not os.path.exists(gj):
        return None
    data = json.load(open(gj, encoding='utf-8'))
    report = []
    yu_lines = 0; split_lines = 0; total_lines = 0
    for pi, pg in enumerate(data['pages']):
        g = [x for x in pg['glyphs'] if x['char'].strip()]
        g.sort(key=lambda c: (c['x'],))  # within band sort later
        # band by baseline: group glyphs whose baseline within 6pt and overlapping x-progression
        # simpler: sort by baseline, walk; new band if baseline jumps >6 from band MIN
        g.sort(key=lambda c: c['baseline'])
        bands = []; cur = []
        for x in g:
            if cur and (x['baseline'] - min(c['baseline'] for c in cur)) > 6:
                bands.append(cur); cur = []
            cur.append(x)
        if cur: bands.append(cur)
        for band in bands:
            total_lines += 1
            bys = collections.defaultdict(list)
            for c in band:
                bys[round(c['baseline'] * 2) / 2].append(c)
            has_yu = any(font_tag(c['font_family']) in ('YuMincho', 'YuGothic') for c in band)
            if has_yu: yu_lines += 1
            distinct = sorted(bys)
            if len(distinct) > 1 and (distinct[-1] - distinct[0]) > 1.5:
                split_lines += 1
                # fonts per sub-baseline
                desc = []
                for bl in distinct:
                    fams = collections.Counter(font_tag(c['font_family']) for c in bys[bl])
                    desc.append('%.1f:%s(%d)' % (bl, fams.most_common(1)[0][0], len(bys[bl])))
                txt = ''.join(c['char'] for c in sorted(band, key=lambda c: c['x']))[:24]
                if len(report) < 14:
                    report.append('  p%d split %.1fpt %s | %r' % (pi + 1, distinct[-1] - distinct[0], ' '.join(desc), txt))
    return total_lines, yu_lines, split_lines, report

def main():
    docs = sys.argv[1:]
    L = ['S516 within-line baseline-split scan (mixed-font Yu Mincho renderer bug)']
    for d in docs:
        path = d if os.path.isabs(d) else os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx', d)
        r = scan(path)
        if r is None:
            L.append('%s: DUMP FAILED' % os.path.basename(d)); continue
        tot, yu, split, rep = r
        L.append('')
        L.append('=== %s : %d lines, %d Yu-lines, %d SPLIT lines' % (os.path.basename(d), tot, yu, split))
        L.extend(rep)
    txt = '\n'.join(L)
    io.open('c:/tmp/_s516_split.txt', 'w', encoding='utf-8').write(txt + '\n')
    for line in txt.split('\n'):
        try: print(line)
        except Exception: print(line.encode('ascii', 'replace').decode())

if __name__ == '__main__':
    main()
