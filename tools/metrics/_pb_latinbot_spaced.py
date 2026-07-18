# -*- coding: utf-8 -*-
"""Latin no-grid page-bottom + baseline placement for SPACED (auto-multiple)
lines — the S937 derivation (2026-07-19).

S827 pinned the SINGLE-spacing threshold = the full hhea box (TNR-12,
lineGap 87/2048 = degenerate); _pb_latinbot_cal confirmed the box rule for
Calibri-11 single. policies__001cf65cd72d881c exposes the SPACED case:
line=259 Calibri-12 (spaced 15.808, natural hhea 14.648, extra 1.16pt,
lineGap 452/2048 = 2.65pt @12) — Word keeps a line whose baseline is
765.94 under cbot 769.92 (room 3.98pt below the baseline), while Oxi's
cursor+nat_hhea test demands 5.24 below its baseline and pushes.

Two readouts per render:
  1. page-1 kept TARGET count vs a bottom-margin 2tw sweep -> the flip pins
     the capacity rule.
  2. the FIRST line's baseline on page 1 (fitz span origin) minus the top
     margin -> where Word places a spaced line's baseline at a page top
     (extra leading above / below / split), i.e. the cursor->baseline
     convention Oxi must match.

Configs: line=259 (the discriminating spaced case) and line=240 (control,
must reproduce the S827/S835 box rule).

Usage: python _pb_latinbot_spaced.py gen | measure | read
"""
import os, sys, glob, shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_latinbot_spaced")


def para(i, line, target=False):
    r = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/>'
    sp = f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="auto"/>'
    txt = f'TARGET{i:02d} marker line.' if target else f'Filler line {i:02d} for the ladder.'
    return (f'<w:p><w:pPr>{sp}<w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{txt}</w:t></w:r></w:p>')


def build(line, n, bottom):
    pgsz = '<w:pgSz w:w="12240" w:h="15840"/>'
    mar = (f'<w:pgMar w:top="1440" w:right="1440" w:bottom="{bottom}" '
           f'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>')
    body = ''.join(para(i, line) for i in range(n))
    body += ''.join(para(i, line, target=True) for i in range(8))
    body += pg.sectpr(pgsz=pgsz, mar=mar, grid='')
    return pg.doc(body)


# Letter 792pt, top 72. line=259: pitch 15.808 -> 40 fillers end at 704.3.
# line=240: pitch 14.648 -> 43 fillers end at 701.9.
CASES = [('259', 259, 40, list(range(400, 561, 8))),
         ('240', 240, 43, list(range(400, 561, 8)))]


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    n = 0
    for tag, line, k, bots in CASES:
        for b in bots:
            pg.write_docx(os.path.join(OUTDIR, f'lbs_{tag}_b{b:04d}.docx'),
                          build(line, k, b), font='Calibri', sz='24',
                          compat='15', cpunct=False)
            n += 1
    print('generated', n)


def measure():
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for f in sorted(glob.glob(os.path.join(OUTDIR, '*.docx'))):
            pdf = f[:-5] + '.pdf'
            if os.path.exists(pdf):
                continue
            tmp = f[:-5] + '_t.docx'
            shutil.copy(f, tmp)
            doc = word.Documents.Open(os.path.abspath(tmp), ReadOnly=True)
            doc.ExportAsFixedFormat(os.path.abspath(pdf), 17)
            doc.Close(False)
            os.remove(tmp)
    finally:
        word.Quit()
    read()


def read():
    import fitz
    for f in sorted(glob.glob(os.path.join(OUTDIR, '*.pdf'))):
        pdf = fitz.open(f)
        p1 = pdf[0].get_text()
        n_t = sum(1 for i in range(8) if f'TARGET{i:02d}' in p1)
        first_base = None
        deepest = None
        for b in pdf[0].get_text('dict')['blocks']:
            for l in b.get('lines', []):
                t = ''.join(s['text'] for s in l['spans']).strip()
                if not t:
                    continue
                y = l['spans'][0]['origin'][1]
                if first_base is None or y < first_base:
                    first_base = y
                if deepest is None or y > deepest:
                    deepest = y
        print(f'{os.path.basename(f)} p1_targets={n_t} first_base={first_base:.2f} deepest_base={deepest:.2f}')


if __name__ == '__main__':
    {'gen': gen, 'measure': measure, 'read': read}[sys.argv[1]]()
