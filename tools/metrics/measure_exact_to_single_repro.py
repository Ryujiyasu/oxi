"""S107 measure: Word vs Oxi y-pos of each paragraph in V1-V5 repros."""
import json, os, subprocess, sys, tempfile
from pathlib import Path
import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path('c:/Users/ryuji/oxi-main')
REPRO_DIR = ROOT / 'tools/metrics/exact_to_single_repro'
RENDERER = ROOT / 'tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe'


def measure_word(docx, word):
    doc = word.Documents.Open(str(docx.absolute()), ReadOnly=True)
    try:
        rows = []
        for i, p in enumerate(doc.Paragraphs):
            rng = p.Range
            collapsed = doc.Range(rng.Start, rng.Start)
            y = collapsed.Information(6)
            txt = rng.Text.rstrip('\r\n')
            rows.append({'idx': i, 'y_pg': y, 'text': txt[:30] if txt else '<empty>'})
        return rows
    finally:
        doc.Close(SaveChanges=False)


def measure_oxi(docx):
    with tempfile.TemporaryDirectory() as tmp:
        prefix = os.path.join(tmp, 'p_')
        dump = os.path.join(tmp, 'layout.json')
        proc = subprocess.run([str(RENDERER), str(docx), prefix, '--dump-layout=' + dump],
                              capture_output=True, text=True, timeout=60)
        if proc.returncode != 0:
            print("  ERR:", proc.stderr[:300])
            return []
        with open(dump, encoding='utf-8') as f:
            d = json.load(f)
    page = d.get('pages', [{}])[0]
    paras = {}
    for el in page.get('elements', []):
        if el.get('type') != 'text': continue
        pi = el.get('para_idx')
        if pi is None: continue
        if pi not in paras:
            paras[pi] = {'pi': pi, 'y': el['y'], 'text': ''}
        paras[pi]['y'] = min(paras[pi]['y'], el['y'])
        paras[pi]['text'] += el.get('text', '')
    return sorted(paras.values(), key=lambda r: r['pi'])


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        for docx in sorted(REPRO_DIR.glob('*.docx')):
            print(f"\n=== {docx.name} ===")
            try:
                w_rows = measure_word(docx, word)
                o_rows = measure_oxi(docx)
                print(f"  {'idx':>3}  {'W y':>8}  {'O y':>8}  {'W-O':>6}  {'W gap':>6}  {'O gap':>6}  text")
                prev_w = None
                prev_o = None
                for i in range(max(len(w_rows), len(o_rows))):
                    w = w_rows[i] if i < len(w_rows) else None
                    o = o_rows[i] if i < len(o_rows) else None
                    wy = w['y_pg'] if w else None
                    oy = o['y'] if o else None
                    diff = (wy - oy) if (wy is not None and oy is not None) else None
                    wgap = (wy - prev_w) if (wy is not None and prev_w is not None) else None
                    ogap = (oy - prev_o) if (oy is not None and prev_o is not None) else None
                    txt = (w['text'] if w else o['text']) if (w or o) else ''
                    fmt = lambda v: f"{v:7.2f}" if v is not None else "    --"
                    fmt6 = lambda v: f"{v:+6.2f}" if v is not None else "    --"
                    print(f"  {i:>3}  {fmt(wy)}  {fmt(oy)}  {fmt6(diff)}  {fmt6(wgap)}  {fmt6(ogap)}  {txt}")
                    prev_w = wy
                    prev_o = oy
            except Exception as e:
                print(f"  ERR: {e}")
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    main()
