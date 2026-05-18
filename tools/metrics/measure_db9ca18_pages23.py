"""S107 follow-up: measure db9ca18 pages 2,3 paragraph positions (Word vs Oxi).

Goal: understand why db9ca18 p.2/p.3 regressed -0.05/-0.08 after the
half-leading fix. Check actual paragraph Y values now to find the
discrepancy.
"""
import json, os, subprocess, sys, tempfile
from pathlib import Path
import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path('c:/Users/ryuji/oxi-main')
DOCX = ROOT / 'tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx'
RENDERER = ROOT / 'tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe'


def measure_word():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    rows = []
    try:
        doc = word.Documents.Open(str(DOCX.absolute()), ReadOnly=True)
        try:
            for i, p in enumerate(doc.Paragraphs):
                rng = p.Range
                start = rng.Start
                collapsed = doc.Range(start, start)
                page = collapsed.Information(3)
                y = collapsed.Information(6)
                txt = rng.Text.rstrip('\r\n')
                rows.append({'idx': i, 'page': page, 'y': y, 'text': txt[:40]})
        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()
    return rows


def measure_oxi():
    with tempfile.TemporaryDirectory() as tmp:
        prefix = os.path.join(tmp, 'p_')
        dump = os.path.join(tmp, 'layout.json')
        proc = subprocess.run([str(RENDERER), str(DOCX), prefix, '--dump-layout=' + dump],
                              capture_output=True, text=True, timeout=300)
        if proc.returncode != 0:
            return None
        with open(dump, encoding='utf-8') as f:
            return json.load(f)


def oxi_para_summary(d):
    paras = {}
    for pi, page in enumerate(d.get('pages', [])):
        for el in page.get('elements', []):
            if el.get('type') != 'text': continue
            pidx = el.get('para_idx')
            if pidx is None: continue
            key = (pi, pidx)
            if key not in paras:
                paras[key] = {'page': pi, 'para_idx': pidx, 'y': el['y'], 'text': ''}
            paras[key]['y'] = min(paras[key]['y'], el['y'])
            paras[key]['text'] += el.get('text', '')
    return paras


def main():
    print("Measuring Word...", flush=True)
    word_rows = measure_word()
    print("Measuring Oxi...", flush=True)
    oxi = measure_oxi()
    oxi_paras = oxi_para_summary(oxi)
    oxi_sorted = sorted(oxi_paras.values(), key=lambda r: (r['page'], r['y']))

    # Pages 2-3 (1-indexed Word = 1-2 0-indexed Oxi)
    print(f"\n=== Word page 2 ===")
    for r in word_rows:
        if r['page'] == 2:
            print(f"  W[{r['idx']:3d}] y={r['y']:7.2f} {r['text']!r}")
    print(f"\n=== Word page 3 ===")
    for r in word_rows:
        if r['page'] == 3:
            print(f"  W[{r['idx']:3d}] y={r['y']:7.2f} {r['text']!r}")
    print(f"\n=== Oxi page 1 (= page 2 in Word) ===")
    for r in oxi_sorted:
        if r['page'] == 1:
            print(f"  O[pi={r['para_idx']:3d}] y={r['y']:7.2f} {r['text'][:40]!r}")
    print(f"\n=== Oxi page 2 (= page 3 in Word) ===")
    for r in oxi_sorted:
        if r['page'] == 2:
            print(f"  O[pi={r['para_idx']:3d}] y={r['y']:7.2f} {r['text'][:40]!r}")


if __name__ == '__main__':
    main()
