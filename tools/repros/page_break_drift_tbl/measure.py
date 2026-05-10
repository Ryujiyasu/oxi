"""Measure PB_DRIFT_TBL — body→table→body cursor advance."""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, '..', '..', '..'))
GDI_EXE = os.path.join(ROOT, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

VARIANTS = [
    "DR_TBL_01", "DR_TBL_02", "DR_TBL_03", "DR_TBL_04", "DR_TBL_05",
    "DR_TBL_06", "DR_TBL_07", "DR_TBL_08", "DR_TBL_09", "DR_TBL_10",
]


def measure_word(docx_path):
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False; word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    try:
        ys = []
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            t = (p.Range.Text or '').strip()
            rng = p.Range
            cr = d.Range(rng.Start, rng.Start)
            ys.append({
                'idx': i, 'text': t[:8],
                'page': int(cr.Information(3)),
                'y': round(cr.Information(6), 3),
                'in_table': bool(cr.Information(12)),  # wdWithInTable=12
            })
        return ys
    finally:
        d.Close(False); word.Quit()


def measure_oxi(docx_path):
    layout_path = os.path.join(HERE, '_tmp_layout.json')
    out_prefix = os.path.join(HERE, '_tmp_oxi')
    cmd = [GDI_EXE, docx_path, out_prefix, '150', f'--dump-layout={layout_path}']
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        print(f'  GDI fail: {r.stderr[:200]}'); return []
    with open(layout_path, encoding='utf-8') as f:
        layout = json.load(f)
    by_pi = {}
    for page_idx, page in enumerate(layout.get('pages', [])):
        for el in page.get('elements', []):
            if el.get('type') != 'text': continue
            pi = el.get('para_idx')
            if pi is None: continue
            text = el.get('text', '')
            y = el.get('y', 0)
            if pi not in by_pi or y < by_pi[pi]['y']:
                by_pi[pi] = {'pi': pi, 'page': page_idx + 1, 'y': y, 'text': text[:8]}
    return [by_pi[k] for k in sorted(by_pi)]


def main():
    results = []
    for vid in VARIANTS:
        path = os.path.join(HERE, f'{vid}.docx')
        if not os.path.exists(path): continue
        print(f'\n=== {vid} ===')
        wd = measure_word(path)
        ox = measure_oxi(path)
        print(f'  Word ({len(wd)} paras):')
        for w in wd[:8]:
            print(f'    idx={w["idx"]:>2} pg={w["page"]} y={w["y"]:>7.2f} in_tbl={w["in_table"]} text={w["text"]!r}')
        print(f'  Oxi ({len(ox)} paras):')
        for o in ox[:8]:
            print(f'    pi={o["pi"]:>2} pg={o["page"]} y={o["y"]:>7.2f} text={o["text"]!r}')

        # Identify body A, table cell para, body C in Word
        # body A = first para, body C = last para; cell paras are in between
        if len(wd) < 3 or len(ox) < 3: continue
        wA, wC = wd[0], wd[-1]
        oA, oC = ox[0], ox[-1]

        print(f'  Body A: Word y={wA["y"]:6.2f} / Oxi y={oA["y"]:6.2f} → init_dy={oA["y"]-wA["y"]:+.2f}')
        print(f'  Body C: Word y={wC["y"]:6.2f} / Oxi y={oC["y"]:6.2f} → final_dy={oC["y"]-wC["y"]:+.2f}')
        through_drift = (oC['y'] - oA['y']) - (wC['y'] - wA['y'])
        print(f'  Total advance through table: Word={wC["y"]-wA["y"]:6.2f} Oxi={oC["y"]-oA["y"]:6.2f} Δ={through_drift:+.2f}')

        results.append({
            'variant': vid,
            'word_paras': wd, 'oxi_paras': ox,
            'body_A_word_y': wA['y'], 'body_A_oxi_y': oA['y'],
            'body_C_word_y': wC['y'], 'body_C_oxi_y': oC['y'],
            'init_dy': round(oA['y'] - wA['y'], 3),
            'final_dy': round(oC['y'] - wC['y'], 3),
            'through_drift': round(through_drift, 3),
        })

    out = os.path.join(HERE, 'measurements.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'\nSaved: {out}')

    print(f'\n{"variant":<11} {"init_dy":>8} {"final_dy":>9} {"through":>8}')
    for r in results:
        print(f'  {r["variant"]:<11} {r["init_dy"]:>+8.2f} {r["final_dy"]:>+9.2f} {r["through_drift"]:>+8.2f}')


if __name__ == '__main__':
    main()
