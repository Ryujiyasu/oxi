"""S107: Word vs Oxi page-2 ALL paragraph y comparison (including empty paras).

Goal: find which inter-paragraph gap differs between Word and Oxi to explain
the +1 paragraph bug on page 2.
"""
import json, os, subprocess, sys, tempfile
from pathlib import Path
import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path('c:/Users/ryuji/oxi-main')
DOCX = ROOT / 'tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx'
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
                y_pg = collapsed.Information(6)  # vertical position relative to page (pt)
                txt = rng.Text.rstrip('\r\n')
                rows.append({
                    'idx': i,
                    'page': page,
                    'y_pg': y_pg,
                    'text': txt,
                    'len': len(txt),
                })
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
            print("oxi failed:", proc.stderr[:500])
            return None
        with open(dump, encoding='utf-8') as f:
            return json.load(f)


def oxi_para_summary(d):
    """For each (page, para_idx), return min y + concatenated text."""
    paras = {}
    for pi, page in enumerate(d.get('pages', [])):
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            pidx = el.get('para_idx')
            if pidx is None:
                continue
            key = (pi, pidx)
            if key not in paras:
                paras[key] = {'page': pi, 'para_idx': pidx, 'y_min': el['y'], 'text': ''}
            paras[key]['y_min'] = min(paras[key]['y_min'], el['y'])
            paras[key]['text'] += el.get('text', '')
    # Empty paragraphs may not appear as text; look for placeholder elements
    return paras


def main():
    print("Measuring Word...", flush=True)
    word_paras = measure_word()
    print(f"  Word total paragraphs: {len(word_paras)}", flush=True)

    print("\nMeasuring Oxi...", flush=True)
    oxi = measure_oxi()
    if oxi is None:
        return
    oxi_paras = oxi_para_summary(oxi)
    print(f"  Oxi total paragraphs: {len(oxi_paras)}", flush=True)

    # Filter Word page 2 (and one before/after for context)
    print(f"\n=== Word: paragraphs on pages 1 (end), 2, 3 (start) ===")
    for p in word_paras:
        if p['page'] in (1, 2, 3):
            txt = p['text'][:30].replace('\n', '⏎')
            txt_marker = '<empty>' if not p['text'].strip() else txt
            print(f"  W[{p['idx']:3d}] pg={p['page']} y={p['y_pg']:7.2f}  len={p['len']:3d}  {txt_marker}")

    # Oxi: page 1 (last few) and page 2 entries
    print(f"\n=== Oxi: paragraphs on pages 0 (idx 0-based), 1, 2 ===")
    oxi_sorted = sorted(oxi_paras.values(), key=lambda r: (r['page'], r['y_min']))
    for r in oxi_sorted:
        if r['page'] in (0, 1, 2):
            txt = r['text'][:30].replace('\n', '⏎')
            txt_marker = '<empty>' if not r['text'].strip() else txt
            print(f"  O[pg={r['page']} pi={r['para_idx']:3d}] y={r['y_min']:7.2f}  {txt_marker}")

    # Compute gaps
    print(f"\n=== Word page-2 paragraphs with gaps ===")
    w_p2 = [p for p in word_paras if p['page'] == 2]
    prev_y = None
    for p in w_p2:
        gap = p['y_pg'] - prev_y if prev_y is not None else 0.0
        prev_y = p['y_pg']
        txt = p['text'][:35].replace('\n', '⏎')
        txt_marker = '<empty>' if not p['text'].strip() else txt
        print(f"  W[{p['idx']:3d}] y={p['y_pg']:7.2f}  gap={gap:+6.2f}  {txt_marker}")

    print(f"\n=== Oxi page-1 (=2nd page) paragraphs with gaps ===")
    o_p2 = [r for r in oxi_sorted if r['page'] == 1]
    prev_y = None
    for r in o_p2:
        gap = r['y_min'] - prev_y if prev_y is not None else 0.0
        prev_y = r['y_min']
        txt = r['text'][:35].replace('\n', '⏎')
        txt_marker = '<empty>' if not r['text'].strip() else txt
        print(f"  O[pi={r['para_idx']:3d}] y={r['y_min']:7.2f}  gap={gap:+6.2f}  {txt_marker}")

    # save full json for analysis
    out = ROOT / 'tools/metrics/d77a_p2_all_paras.json'
    out.write_text(json.dumps({
        'word': word_paras,
        'oxi': list(oxi_paras.values()),
    }, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f"\nSaved {out}")


if __name__ == '__main__':
    main()
