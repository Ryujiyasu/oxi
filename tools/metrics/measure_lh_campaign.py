"""Day 33 part 26 (2026-05-11) — Word line-height campaign measurement.

For each repro in tools/golden-test/repros/lh_campaign/, measure Word's
per-line advance (y of para 2 - y of para 1) via COM. Output JSON
lookup table for future fix.
"""
from __future__ import annotations
import os, sys, json, glob, re
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

REPRO_DIR = 'tools/golden-test/repros/lh_campaign'
OUT_PATH = 'pipeline_data/lh_campaign.json'


def measure(docx_path):
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    abs_path = os.path.abspath(docx_path)
    d = word.Documents.Open(abs_path, ReadOnly=True)
    try:
        n = d.Paragraphs.Count
        ys = []
        for i in range(1, min(n, 5) + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr = d.Range(r.Start, r.Start)
            try:
                pg = int(cr.Information(3))
                y = round(cr.Information(6), 2)
                if pg != 1: break  # only page 1
                ys.append(y)
            except: continue
        # advance = avg pairwise diff
        advances = [ys[i+1] - ys[i] for i in range(len(ys)-1)] if len(ys) >= 2 else []
        return ys, advances
    finally:
        d.Close(False)
        word.Quit()


def parse_name(name):
    """Extract structured fields from filename."""
    # e.g. LH_msMincho_10p5_g360 → {font:'MS Mincho', fs:10.5, grid:'lAC360'}
    m = re.match(r'LH_(.+)', name)
    if not m: return {}
    fields = m.group(1).split('_')
    return {'raw_name': name, 'fields': fields}


def main():
    repros = sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx')))
    print(f'Found {len(repros)} repros')
    results = []
    for r in repros:
        name = os.path.splitext(os.path.basename(r))[0]
        try:
            ys, advances = measure(r)
        except Exception as e:
            print(f'  {name}: ERROR {e}')
            results.append({'name': name, 'error': str(e)[:200]})
            continue
        if not advances:
            print(f'  {name}: ys={ys} advances=[]')
            results.append({'name': name, 'ys': ys, 'advances': [], 'avg_advance': None})
            continue
        avg = round(sum(advances) / len(advances), 3)
        unique = sorted(set(round(a, 2) for a in advances))
        print(f'  {name}: ys={ys} advances={advances} avg={avg} unique={unique}')
        results.append({
            'name': name, 'ys': ys, 'advances': advances,
            'avg_advance': avg, 'unique_advances': unique,
        })
    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        json.dump({'measurements': results}, f, ensure_ascii=False, indent=2)
    print(f'Wrote {OUT_PATH}')


if __name__ == '__main__':
    main()
