"""a1d6 の全段落 Y を COM で測定 (d4d126_drift_origin.py の a1d6 版).

Oxi vs Word の per-paragraph drift パターンを比較するため。
"""
import json, sys, os
import win32com.client

DOCX = os.path.abspath('tools/golden-test/documents/docx/a1d6e4efa2e7_tokumei_08_01-4.docx')
OUT  = os.path.abspath('pipeline_data/a1d6_drift_origin.json')

def main():
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        n = doc.Paragraphs.Count
        print(f'Total paragraphs: {n}')
        rows = []
        for wi in range(1, n + 1):
            p = doc.Paragraphs(wi)
            rng = p.Range
            start_rng = doc.Range(rng.Start, rng.Start)
            try:
                page = int(start_rng.Information(3))
                y    = float(start_rng.Information(6))
                x    = float(start_rng.Information(5))
            except Exception:
                continue
            ls   = float(p.LineSpacing or 0)
            lsr  = int(p.LineSpacingRule or 0)
            sb   = float(p.SpaceBefore or 0)
            in_tbl = bool(rng.Information(12))
            txt  = (rng.Text or '').replace('\r','').replace('\x07','')[:60]
            rows.append(dict(wi=wi, page=page, x=x, y=y,
                             ls=ls, lsr=lsr, sb=sb,
                             in_tbl=in_tbl, text=txt))
            if wi % 100 == 0:
                print(f'  done {wi}/{n}')
        with open(OUT, 'w', encoding='utf-8') as f:
            json.dump({'doc': os.path.basename(DOCX), 'rows': rows}, f, ensure_ascii=False, indent=2)
        print(f'Wrote {OUT}')
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    main()
