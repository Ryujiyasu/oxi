"""S181: COM-measure each RBM variant's per-row y in Word + Oxi,
derive Word's row-height border-allocation rule.

For each variant:
  - Word COM: Document.Paragraphs(i).Range collapsed start Information(6)
    for each paragraph (each row's paragraph)
  - Oxi: --dump-layout, find each row's paragraph y; also OXI_DUMP_TABLE
    to capture row_height value
  - Compute per-row dy (Oxi - Word), per-row Word advance, per-row Oxi
    advance

Output:
  pipeline_data/row_border_matrix_results.json
  stdout: comparison table

Pre-req: Word installed, oxi-gdi-renderer built.

Run:
  python tools/metrics/measure_row_border_matrix.py
"""
from __future__ import annotations
import os, sys, subprocess, json, time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent.parent
DOCS = REPO / 'tools' / 'golden-test' / 'repros' / 'row_border_matrix'
RENDERER = REPO / 'tools' / 'oxi-gdi-renderer' / 'target' / 'release' / 'oxi-gdi-renderer.exe'
OUT_JSON = REPO / 'pipeline_data' / 'row_border_matrix_results.json'

sys.stdout.reconfigure(encoding='utf-8', errors='replace')


def measure_word(docx_path: Path) -> list[dict]:
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    try:
        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        time.sleep(0.5)
        try:
            n_paras = doc.Paragraphs.Count
            out = []
            for pi in range(1, n_paras + 1):
                p = doc.Paragraphs(pi)
                rng = p.Range
                # R30: collapsed start
                cs = doc.Range(rng.Start, rng.Start)
                try:
                    pg = int(cs.Information(3))  # wdActiveEndPageNumber
                    y = float(cs.Information(6))  # wdVerticalPositionRelativeToPage
                except Exception as e:
                    pg, y = None, None
                txt = (rng.Text or '').rstrip('\r\x07')
                out.append({'i': pi, 'page': pg, 'y': y, 'text': txt[:30]})
            return out
        finally:
            doc.Close(False)
    finally:
        word.Quit()


def measure_oxi(docx_path: Path) -> tuple[list[dict], str]:
    """Returns (paragraph_records, tbl_dump_text)."""
    import tempfile
    with tempfile.TemporaryDirectory(prefix='rbm_') as tmp:
        out_prefix = os.path.join(tmp, 'p_')
        dump_path = os.path.join(tmp, 'layout.json')
        env = os.environ.copy()
        env['OXI_DUMP_TABLE'] = '1'
        proc = subprocess.run(
            [str(RENDERER), str(docx_path), out_prefix,
             '--exclude=text,border,shading,box,image,clip',
             f'--dump-layout={dump_path}'],
            capture_output=True, text=True, timeout=60, env=env,
        )
        if proc.returncode != 0:
            raise RuntimeError(f'renderer failed: {proc.stderr[:300]}')
        with open(dump_path, encoding='utf-8') as f:
            dump = json.load(f)
    # Extract paragraph records: for each (page, para_idx, cell_row_idx, cell_col_idx, cell_para_idx)
    # use first occurrence by (y, x)
    out = []
    seen = set()
    for page in dump.get('pages', []):
        pn = page['page']
        texts = [e for e in page.get('elements', []) if e.get('type') == 'text']
        texts.sort(key=lambda e: (e['y'], e['x']))
        for e in texts:
            key = (e.get('para_idx'), e.get('cell_row_idx'),
                   e.get('cell_col_idx'), e.get('cell_para_idx'))
            if key in seen:
                continue
            seen.add(key)
            out.append({
                'key': list(key),
                'page': pn,
                'y': round(e['y'], 3),
                'x': round(e['x'], 3),
                'text': e.get('text', '')[:20],
            })
    return out, proc.stderr


def main():
    if not DOCS.exists():
        print(f'no docs at {DOCS}')
        return
    if not RENDERER.exists():
        print(f'no renderer at {RENDERER}')
        return

    docs = sorted(DOCS.glob('RBM_*.docx'))
    results = []
    for d in docs:
        label = d.stem[4:]  # strip "RBM_"
        print(f'\n=== {label} ===')
        try:
            w = measure_word(d)
        except Exception as e:
            print(f'  Word fail: {e}')
            continue
        # Filter only non-empty text paragraphs (the row content)
        w_rows = [p for p in w if p['text'].strip()]
        try:
            o, tbl_dump = measure_oxi(d)
        except Exception as e:
            print(f'  Oxi fail: {e}')
            continue
        o_rows = [p for p in o if p['text'].strip()]
        # Align positionally (assumes same number of non-empty paragraphs)
        n = min(len(w_rows), len(o_rows))
        pairs = []
        for i in range(n):
            wp = w_rows[i]
            op = o_rows[i]
            pair = {
                'i': i,
                'w_y': wp['y'], 'o_y': op['y'],
                'w_page': wp['page'], 'o_page': op['page'],
                'w_text': wp['text'], 'o_text': op['text'],
                'o_key': op['key'],
            }
            if wp['y'] is not None and op['y'] is not None and wp['page'] == op['page']:
                pair['dy'] = round(op['y'] - wp['y'], 3)
            else:
                pair['dy'] = None
            pairs.append(pair)

        # Per-row Word advance, Oxi advance
        for i in range(1, len(pairs)):
            prev = pairs[i-1]
            cur = pairs[i]
            if prev['w_page'] == cur['w_page']:
                cur['w_step'] = round(cur['w_y'] - prev['w_y'], 3)
            if prev['o_page'] == cur['o_page']:
                cur['o_step'] = round(cur['o_y'] - prev['o_y'], 3)

        # Extract row_h from tbl_dump
        row_h_values = []
        for line in tbl_dump.split('\n'):
            if 'pre_correction row_height=' in line:
                # Format: [TBL_DUMP] row=N pre_correction row_height=X.YYY max_actual_cell_h=Z.ZZZ
                try:
                    parts = line.split()
                    for p in parts:
                        if p.startswith('row_height='):
                            row_h_values.append(float(p.split('=')[1]))
                except Exception:
                    pass

        result = {
            'label': label,
            'n_pairs': len(pairs),
            'pairs': pairs,
            'oxi_row_h_values': row_h_values,
        }
        results.append(result)

        # Print summary
        print(f'  n_word_text={len(w_rows)} n_oxi_text={len(o_rows)}')
        for p in pairs:
            ws = p.get('w_step', '-')
            os_ = p.get('o_step', '-')
            ws_str = f'{ws:>+7.2f}' if isinstance(ws, float) else f'{ws:>7}'
            os_str = f'{os_:>+7.2f}' if isinstance(os_, float) else f'{os_:>7}'
            dy = p.get('dy', '-')
            dy_str = f'{dy:>+7.2f}' if isinstance(dy, float) else f'{dy:>7}'
            print(f'    i={p["i"]:>2} w_y={p["w_y"]:>6.2f} o_y={p["o_y"]:>6.2f} dy={dy_str} w_step={ws_str} o_step={os_str} text={p["w_text"][:15]!r}')
        print(f'  Oxi row_h values: {row_h_values}')

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'\nResults -> {OUT_JSON}')


if __name__ == '__main__':
    main()
