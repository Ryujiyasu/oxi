# -*- coding: utf-8 -*-
"""Render-truth glyph matcher, PDF variant (for table docs where EMF export is empty).
Word glyphs from word_pdf_glyphs.py (per-char origin x,y in pt, page-absolute).
Oxi glyphs from the layout dump (elements may be MULTI-char -> expand to per-char at
x + k*(w/n), uniform CJK pitch). Align char streams with difflib; calibrate by median
(both page-absolute so cal is small). Report rdx/rdy residual + per-line profile.
Usage: rendertruth_match_pdf.py <word_pdf.json> <oxi_dump.json> <page_idx0> <out.json>
cp932-safe: UTF-8 file, JSON out, ASCII verdicts."""
import json, sys, difflib, statistics


def cluster_lines(glyphs, tol):
    gs = sorted(glyphs, key=lambda g: (round(g['y'], 1), g['x']))
    lines = []
    for g in gs:
        if lines and abs(g['y'] - lines[-1]['y']) < tol:
            lines[-1]['gs'].append(g); lines[-1]['y'] = lines[-1]['gs'][0]['y']
        else:
            lines.append({'y': g['y'], 'gs': [g]})
    for L in lines:
        L['gs'].sort(key=lambda g: g['x'])
    return lines


def main():
    wpath, dump_path, pidx, out = sys.argv[1], sys.argv[2], int(sys.argv[3]), sys.argv[4]
    W = json.load(open(wpath, encoding='utf-8'))['pages'][pidx]['glyphs']
    O = json.load(open(dump_path, encoding='utf-8'))['pages'][pidx]['elements']

    wg = [{'char': g['char'], 'x': g['x'], 'y': g['y']} for g in W if g['char'].strip()]

    # Oxi: expand multi-char elements to per-char using uniform sub-pitch w/n.
    og = []
    for e in O:
        if e.get('type') != 'text':
            continue
        t = e.get('text', '')
        if not t.strip():
            continue
        n = len(t)
        w = e.get('w', 0) or 0
        pitch = (w / n) if n else 0
        for k, ch in enumerate(t):
            if not ch.strip():
                continue
            og.append({'char': ch, 'x': e['x'] + k * pitch, 'y': e['y'],
                       'fs': e.get('font_size', 0),
                       'cell': (e.get('cell_row_idx'), e.get('cell_col_idx'))})

    wlines = cluster_lines(wg, 5.0)
    olines = cluster_lines(og, 5.0)
    wseq = [g for L in wlines for g in L['gs']]
    oseq = [g for L in olines for g in L['gs']]
    wchars = [g['char'] for g in wseq]
    ochars = [g['char'] for g in oseq]

    sm = difflib.SequenceMatcher(None, wchars, ochars, autojunk=False)
    matched = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            for dd in range(i2 - i1):
                w_, o_ = wseq[i1 + dd], oseq[j1 + dd]
                matched.append({'char': w_['char'], 'wx': w_['x'], 'wy': w_['y'],
                                'ox': o_['x'], 'oy': o_['y'], 'fs': o_.get('fs', 0),
                                'cell': o_.get('cell')})
    cal_x = statistics.median(m['ox'] - m['wx'] for m in matched)
    cal_y = statistics.median(m['oy'] - m['wy'] for m in matched)
    for m in matched:
        m['rdx'] = round((m['ox'] - m['wx']) - cal_x, 2)
        m['rdy'] = round((m['oy'] - m['wy']) - cal_y, 2)
    rdxs = [m['rdx'] for m in matched]
    rdys = [m['rdy'] for m in matched]
    summ = {'word_glyphs': len(wg), 'oxi_glyphs': len(og),
            'word_lines': len(wlines), 'oxi_lines': len(olines),
            'matched': len(matched), 'match_ratio': round(sm.ratio(), 4),
            'cal_x': round(cal_x, 2), 'cal_y': round(cal_y, 2),
            'rdx_absmean': round(statistics.mean(abs(v) for v in rdxs), 2),
            'rdx_std': round(statistics.pstdev(rdxs), 2),
            'rdx_range': [round(min(rdxs), 1), round(max(rdxs), 1)],
            'rdy_mean': round(statistics.mean(rdys), 2),
            'rdy_absmean': round(statistics.mean(abs(v) for v in rdys), 2),
            'rdy_std': round(statistics.pstdev(rdys), 2),
            'rdy_range': [round(min(rdys), 1), round(max(rdys), 1)]}
    print("=== b35 PDF render-truth match ===")
    for k, v in summ.items():
        print("  %-14s %s" % (k, v))
    print("\nrdy>0 => Oxi BELOW Word (too low); rdy<0 => Oxi ABOVE (too high). vertical cell drift.")
    json.dump({'summary': summ, 'matched': matched}, open(out, 'w', encoding='utf-8'),
              ensure_ascii=False)
    print("wrote", out)


if __name__ == '__main__':
    main()
