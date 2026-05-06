# -*- coding: utf-8 -*-
"""
Test hypothesis: when section has w:footer=N twips but no footer.xml, does
Word still reserve footer space below pgH - footer when computing body floor?

Method:
- For 5 candidate docs, get PageSetup.PageHeight, .BottomMargin, .FooterDistance
- Get last paragraph on each page; record max y of body content via Information(6)
- If max_body_y < pgH - footer_distance for all pages, hypothesis confirmed.
- If max_body_y >= pgH - bottom_margin even sometimes, hypothesis falsified.
"""
import sys, os, json, glob
sys.stdout.reconfigure(encoding='utf-8')

import win32com.client as wc
from pathlib import Path

PT_PER_INCH = 72.0

def points(twips):
    return twips / 20.0

def main():
    docs = [
        'tools/golden-test/documents/docx/1636d28e2c46_tokumei_08_04.docx',
        'tools/golden-test/documents/docx/2ea81a8441cc_0025006-192.docx',
        'tools/golden-test/documents/docx/a47e6c6b2ca1_order_08.docx',
        'tools/golden-test/documents/docx/b5f706e9f6ad_kyodokenkyuyoushiki_bessi.docx',
        'tools/golden-test/documents/docx/e201249db062_tokumei_08_05.docx',
    ]
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    results = []
    for doc_path in docs:
        ap = os.path.abspath(doc_path)
        if not os.path.exists(ap):
            continue
        print(f'\n=== {Path(doc_path).name} ===')
        d = word.Documents.Open(ap, ReadOnly=True)
        try:
            ps = d.PageSetup
            ph = ps.PageHeight
            bm = ps.BottomMargin
            fd = ps.FooterDistance
            print(f'  PageHeight={ph:.2f}pt  BottomMargin={bm:.2f}pt  FooterDistance={fd:.2f}pt')
            print(f'  body_floor_if_only_bm = {ph - bm:.2f}pt')
            print(f'  body_floor_if_footer = {ph - fd:.2f}pt')

            # Per-page max body y (last paragraph y + line height estimate)
            paras = d.Paragraphs
            n = paras.Count
            page_max_y = {}
            for i in range(1, n + 1):
                rng = paras(i).Range
                start = d.Range(rng.Start, rng.Start)
                page = start.Information(3)
                y = start.Information(6)
                if y is None or y < 0:
                    continue
                if page not in page_max_y or y > page_max_y[page]:
                    page_max_y[page] = y
            doc_info = {
                'name': Path(doc_path).name,
                'page_height_pt': ph,
                'bottom_margin_pt': bm,
                'footer_distance_pt': fd,
                'body_floor_bm': ph - bm,
                'body_floor_footer': ph - fd,
                'page_max_y': dict(sorted(page_max_y.items())),
            }
            for pg, y in sorted(page_max_y.items()):
                vs_bm = y < (ph - bm)
                vs_fd = y < (ph - fd)
                print(f'  page {pg}: max_body_y={y:.2f}pt  <bm_floor={vs_bm}  <fd_floor={vs_fd}')
            results.append(doc_info)
        finally:
            d.Close(SaveChanges=False)
    word.Quit()
    out_path = 'pipeline_data/ra_manual_measurements.json'
    try:
        existing = json.load(open(out_path, encoding='utf-8'))
    except Exception:
        existing = []
    existing.append({
        'tag': 'footer_reservation_no_footer_xml',
        'date': '2026-05-06',
        'docs': results,
    })
    json.dump(existing, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f'\nSaved to {out_path}')

if __name__ == '__main__':
    main()
