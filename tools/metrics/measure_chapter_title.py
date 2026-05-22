"""S188: Measure 4-variant chapter-title-line repros (Word COM + Oxi)
to confirm 14pt bold MS Mincho over-allocation.

For each variant: get per-paragraph y from Word + Oxi.
The chapter title is paragraph 6 (1-indexed). The advance from para 5
to para 6 = body line height. The advance from para 6 to para 7 =
chapter title line height (= what Word/Oxi allocate for the chapter
title).

If Oxi advances >7pt MORE than Word for variant A (14pt bold MS) but
NOT for B/C/D, the bug is specific to that font/size/bold tuple.

Run: python tools/metrics/measure_chapter_title.py
"""
from __future__ import annotations
import os, sys, subprocess, json, time, tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent.parent
DOCS = REPO / 'tools' / 'golden-test' / 'repros' / 'chapter_title'
RENDERER = REPO / 'tools' / 'oxi-gdi-renderer' / 'target' / 'release' / 'oxi-gdi-renderer.exe'
OUT = REPO / 'pipeline_data' / 'chapter_title_results.json'

sys.stdout.reconfigure(encoding='utf-8', errors='replace')


def measure_word(path):
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    try:
        doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
        time.sleep(0.5)
        try:
            n = doc.Paragraphs.Count
            out = []
            for i in range(1, n + 1):
                p = doc.Paragraphs(i)
                rng = p.Range
                cs = doc.Range(rng.Start, rng.Start)
                try:
                    pg = int(cs.Information(3))
                    y = float(cs.Information(6))
                except Exception:
                    pg, y = None, None
                txt = (rng.Text or '').rstrip('\r\x07')
                out.append({'i': i, 'page': pg, 'y': y, 'text': txt[:30]})
            return out
        finally:
            doc.Close(False)
    finally:
        word.Quit()


def measure_oxi(path):
    with tempfile.TemporaryDirectory() as tmp:
        out_prefix = os.path.join(tmp, 'p_')
        dump_path = os.path.join(tmp, 'l.json')
        proc = subprocess.run(
            [str(RENDERER), str(path), out_prefix,
             '--exclude=text,border,shading,box,image,clip',
             f'--dump-layout={dump_path}'],
            capture_output=True, text=True, timeout=60,
        )
        if proc.returncode != 0:
            raise RuntimeError(proc.stderr[:200])
        with open(dump_path, encoding='utf-8') as f:
            dump = json.load(f)
    out = []
    seen = set()
    for page in dump.get('pages', []):
        for e in sorted(page['elements'], key=lambda r: (r['y'], r['x'])):
            if e.get('type') != 'text': continue
            key = (e.get('para_idx'), e.get('cell_para_idx'))
            if key in seen: continue
            seen.add(key)
            out.append({
                'pi': e.get('para_idx'),
                'page': page['page'],
                'y': round(e['y'], 3),
                'text_y_off': round(e.get('text_y_off', 0), 3),
                'fs': e.get('font_size'),
                'text': e.get('text', '')[:30],
            })
    return out


def main():
    results = []
    for d in sorted(DOCS.glob('CT_*.docx')):
        label = d.stem[3:]
        print(f'\n=== {label} ===')
        try:
            w = measure_word(d)
        except Exception as e:
            print(f'  Word fail: {e}'); continue
        try:
            o = measure_oxi(d)
        except Exception as e:
            print(f'  Oxi fail: {e}'); continue

        w_text = [p for p in w if p['text'].strip()]
        o_text = [p for p in o if p['text'].strip()]
        n = min(len(w_text), len(o_text))

        # Print first 12 pairs with per-step
        print(f'  n_word={len(w_text)} n_oxi={len(o_text)}')
        print(f'  {"i":>2} {"wy":>6} {"oy":>6} {"dy":>6} {"w_step":>7} {"o_step":>7} text')
        rows = []
        for i in range(n):
            wp = w_text[i]; op = o_text[i]
            dy = round(op['y'] - wp['y'], 2)
            w_step = round(wp['y'] - w_text[i-1]['y'], 2) if i > 0 else None
            o_step = round(op['y'] - o_text[i-1]['y'], 2) if i > 0 else None
            ws = f'{w_step:>+7.2f}' if w_step is not None else '       -'
            os_ = f'{o_step:>+7.2f}' if o_step is not None else '       -'
            print(f'  {i:>2} {wp["y"]:>6.2f} {op["y"]:>6.2f} {dy:>+6.2f} {ws} {os_} {wp["text"][:30]!r}')
            rows.append({'i': i, 'w_y': wp['y'], 'o_y': op['y'], 'dy': dy, 'w_step': w_step, 'o_step': o_step, 'text': wp['text']})
        results.append({'label': label, 'pairs': rows})

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'\nWrote {OUT}')


if __name__ == '__main__':
    main()
