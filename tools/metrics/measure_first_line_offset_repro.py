"""S107: measure first-line Y offset from topMargin for each font/size combo."""
import json, os, subprocess, sys, tempfile
from pathlib import Path
import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path('c:/Users/ryuji/oxi-main')
REPRO_DIR = ROOT / 'tools/metrics/first_line_offset_repro'
RENDERER = ROOT / 'tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe'
TOP_MARGIN_PT = 1418 / 20.0  # 70.9pt


def measure_word(docx, word):
    doc = word.Documents.Open(str(docx.absolute()), ReadOnly=True)
    try:
        p = doc.Paragraphs(1)
        collapsed = doc.Range(p.Range.Start, p.Range.Start)
        return collapsed.Information(6)
    finally:
        doc.Close(SaveChanges=False)


def measure_oxi(docx):
    with tempfile.TemporaryDirectory() as tmp:
        prefix = os.path.join(tmp, 'p_')
        dump = os.path.join(tmp, 'layout.json')
        proc = subprocess.run([str(RENDERER), str(docx), prefix, '--dump-layout=' + dump],
                              capture_output=True, text=True, timeout=60)
        if proc.returncode != 0:
            return None
        with open(dump, encoding='utf-8') as f:
            d = json.load(f)
    page = d.get('pages', [{}])[0]
    min_y = None
    for el in page.get('elements', []):
        if el.get('type') != 'text' or el.get('para_idx') != 0:
            continue
        if min_y is None or el['y'] < min_y:
            min_y = el['y']
    return min_y


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        print(f"  topMargin = {TOP_MARGIN_PT:.2f}pt")
        print(f"  {'doc':<25} {'W y':>8} {'O y':>8} {'W-tm':>7} {'O-tm':>7} {'W-O':>7}")
        for docx in sorted(REPRO_DIR.glob('*.docx')):
            try:
                wy = measure_word(docx, word)
                oy = measure_oxi(docx)
                w_off = wy - TOP_MARGIN_PT
                o_off = oy - TOP_MARGIN_PT if oy else None
                diff = wy - oy if oy else None
                print(f"  {docx.stem:<25} {wy:8.2f} {oy or 0:8.2f} {w_off:+7.2f} {o_off or 0:+7.2f} {diff or 0:+7.2f}")
            except Exception as e:
                print(f"  {docx.stem}: ERR {e}")
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    main()
