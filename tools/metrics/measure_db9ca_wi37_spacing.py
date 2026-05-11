"""Day 33 part 27 (2026-05-11) — Measure db9ca wi=37 paragraph spacing.

Day 33 part 26 falsified Day 33 part 24 lh-mismatch hypothesis: Word
uses same 18pt/line as Oxi for db9ca's grid setup. The 19pt advance
observed is from paragraph spacing (style "表 (緑) 3"), not lh.

This script measures Word's actual Format.SpaceBefore/After for wi=37
and neighbors, to confirm the source of the +1pt/+2pt drift.
"""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOCX = 'tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx'


def main():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    abs_path = os.path.abspath(DOCX)
    d = word.Documents.Open(abs_path, ReadOnly=True)
    try:
        n = d.Paragraphs.Count
        for i in range(33, min(40, n) + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr = d.Range(r.Start, r.Start)
            try: pg = int(cr.Information(3))
            except: pg = -1
            try: y = round(cr.Information(6), 2)
            except: y = -1
            try: text = (r.Text or '').rstrip('\r\n')[:35]
            except: text = ''
            try: sb = round(p.Format.SpaceBefore, 2)
            except: sb = -1
            try: sa = round(p.Format.SpaceAfter, 2)
            except: sa = -1
            try: sba = bool(p.Format.SpaceBeforeAuto)
            except: sba = None
            try: saa = bool(p.Format.SpaceAfterAuto)
            except: saa = None
            try: ls = round(p.Format.LineSpacing, 2)
            except: ls = -1
            try: lsr = p.Format.LineSpacingRule
            except: lsr = -1
            try: sty = str(p.Style.NameLocal)
            except: sty = '?'
            try: snap = p.Format.SnapToGrid
            except: snap = -1
            print(f'wi={i:>3} pg={pg} y={y:>6} | sb={sb} sa={sa} sba={sba} saa={saa} | ls={ls} lsr={lsr} | snap={snap} sty={sty!r} | text={text!r}')
    finally:
        d.Close(False)
        word.Quit()


if __name__ == '__main__':
    main()
