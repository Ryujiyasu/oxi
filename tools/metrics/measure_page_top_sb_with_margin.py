"""Reliable page-top sb measurement: read top_margin per doc, then verify
y_first_on_page = top_margin + sb (APPLY) or y = top_margin (SUPPRESS).

Default top_margin is per-section pgMar/top.
"""
import json, sys, os, zipfile
import xml.etree.ElementTree as ET
import win32com.client

NS = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

DOCS = [
    'tools/golden-test/documents/docx/b5f706e9f6ad_kyodokenkyuyoushiki_bessi.docx',
    'tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx',
]

def get_top_margin(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        with z.open('word/document.xml') as f:
            xml = f.read().decode('utf-8')
    tree = ET.fromstring(xml)
    sectPr = tree.find('.//w:sectPr', NS)
    pgMar = sectPr.find('w:pgMar', NS)
    top_tw = int(pgMar.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top','1440'))
    return top_tw / 20.0

def measure(docx_path):
    top_margin = get_top_margin(docx_path)
    print(f'\n=== {os.path.basename(docx_path)} (top_margin={top_margin:.2f}pt) ===')
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    try:
        n = doc.Paragraphs.Count
        first_on_page = {}
        for wi in range(1, n + 1):
            p = doc.Paragraphs(wi)
            rng = p.Range
            text = (rng.Text or '').replace('\r','').replace('\x07','').strip()
            start_rng = doc.Range(rng.Start, rng.Start)
            try:
                page = int(start_rng.Information(3))
                y    = float(start_rng.Information(6))
            except Exception:
                continue
            if y > 800 or y < 0: continue
            if page not in first_on_page:
                first_on_page[page] = dict(
                    wi=wi, page=page, y=y,
                    sb=float(p.SpaceBefore),
                    text=text[:50],
                )
        rows = [first_on_page[p] for p in sorted(first_on_page.keys())]
        print(f'{"page":>4} {"wi":>4} {"y":>7} {"y-tm":>7} {"sb":>6} {"verdict":>10} text')
        applied = 0
        suppressed = 0
        for r in rows:
            y_off = r['y'] - top_margin
            if r['page'] == 1:
                v = 'p1-skip'
            elif r['sb'] < 0.1:
                v = 'sb=0'
            elif abs(y_off - r['sb']) < 1.5:
                v = 'APPLY'
                applied += 1
            elif y_off < 1.5:
                v = 'SUPPRESS'
                suppressed += 1
            else:
                v = 'AMBIG'
            print(f'{r["page"]:>4} {r["wi"]:>4} {r["y"]:>7.2f} {y_off:>+7.2f} {r["sb"]:>6.2f} {v:>10} {r["text"]!r}')
        print(f'  VERDICT: APPLIED={applied}, SUPPRESSED={suppressed}')
        return rows, applied, suppressed
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    for docx in DOCS:
        if not os.path.exists(docx):
            print(f'MISSING: {docx}')
            continue
        measure(docx)
