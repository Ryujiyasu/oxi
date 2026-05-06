# -*- coding: utf-8 -*-
"""COM-measure where Word breaks pages on the repro_footer_floor_*.docx variants."""
import os, sys, json
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

variants = [
    ('repro_footer_floor_V1.docx', 'bm=5pt fd=72pt   plain'),
    ('repro_footer_floor_V2.docx', 'bm=100pt fd=72pt plain'),
    ('repro_footer_floor_V3.docx', 'bm=25pt fd=36pt  plain'),
    ('repro_footer_floor_V4.docx', 'bm=19.85pt fd=45.35pt titlePg + grid=lines pitch=323  (2ea81a-like)'),
    ('repro_footer_floor_V5.docx', 'bm=19.85pt fd=45.35pt grid=lines pitch=323 (no titlePg)'),
    ('repro_footer_floor_V6.docx', 'bm=19.85pt fd=45.35pt titlePg only (grid=default 312)'),
]

word = wc.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
out = []
for name, label in variants:
    path = os.path.abspath(f'tools/metrics/_repros/{name}')
    print(f'\n=== {name}  {label} ===')
    d = word.Documents.Open(path, ReadOnly=True)
    try:
        ps = d.PageSetup
        ph, bm, fd = ps.PageHeight, ps.BottomMargin, ps.FooterDistance
        print(f'  pgH={ph:.2f}  bm={bm:.2f}  fd={fd:.2f}')
        print(f'  floor_bm={ph-bm:.2f}  floor_fd={ph-fd:.2f}  floor_max={ph-max(bm,fd):.2f}')
        page_max_y = {}
        page_first_y = {}
        for i in range(1, d.Paragraphs.Count + 1):
            rng = d.Paragraphs(i).Range
            start = d.Range(rng.Start, rng.Start)
            page = start.Information(3)
            y = start.Information(6)
            if y is None or y < 0:
                continue
            if page not in page_max_y or y > page_max_y[page]:
                page_max_y[page] = y
            if page not in page_first_y or y < page_first_y[page]:
                page_first_y[page] = y
        for pg in sorted(page_max_y):
            print(f'  page {pg}: first_y={page_first_y[pg]:.2f}  max_y={page_max_y[pg]:.2f}')
        out.append({'name': name, 'pgH': ph, 'bm': bm, 'fd': fd, 'page_max_y': page_max_y, 'page_first_y': page_first_y})
    finally:
        d.Close(SaveChanges=False)
word.Quit()

json.dump(out, open('pipeline_data/footer_floor_repro_results.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
print('\nSaved pipeline_data/footer_floor_repro_results.json')
