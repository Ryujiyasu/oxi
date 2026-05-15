"""COM measurement: find where d4d126's 24.32pt cumulative y-drift originates.

Question: at end of page 6, Oxi's body-para y (pi=49) is 712.18 while Word's
matching wi=336 is at y=736.50 — 24.32pt drift. Where does this accumulate?

Strategy: measure Word page/y for every paragraph wi=1..340. Cross-match with
Oxi's --dump-layout (matched by first 16 chars of text). For each pair compute
oxi_y - word_y as a function of wi (or document position). The drift origin is
where the cumulative delta starts climbing.

Output: pipeline_data/d4d126_drift_origin.json with per-wi (word_page, word_y, text).
"""
import json, sys, os
import win32com.client

DOCX = os.path.abspath('tools/golden-test/documents/docx/d4d126dfe1d9_tokumei_08_01-3.docx')
OUT  = os.path.abspath('pipeline_data/d4d126_drift_origin.json')

def main():
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
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
            page = start_rng.Information(3)
            y    = start_rng.Information(6)
            x    = start_rng.Information(5)
            ls   = p.LineSpacing
            lsr  = p.LineSpacingRule
            sb   = p.SpaceBefore
            txt  = (rng.Text or '').replace('\r','').replace('\x07','')[:60]
            rows.append(dict(wi=wi, page=int(page), x=float(x), y=float(y),
                             ls=float(ls), lsr=int(lsr), sb=float(sb), text=txt))
            if wi % 50 == 0:
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
