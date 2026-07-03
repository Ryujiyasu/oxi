# -*- coding: utf-8 -*-
"""Ruby-in-grid cell-count sweep — derive Word's 2-cell threshold.

One docx per base size: a sequence of ONE-LINE paragraphs on a typed docGrid
(lines, pitch 360), each with a single ruby run of a swept (hps, hpsRaise)
config, separated by no-ruby CONTROL paragraphs. Word's per-paragraph y-gap
(measure_pagination_word.py) then reads the cell count each ruby line gets:
gap 18 = 1 cell, 36 = 2 cells, 54 = 3 cells.

Run: python tools/metrics/_ruby_grid_sweep.py
     -> probervsweep21_rubygridsweep.docx (base sz=21) + probervsweep24 (sz=24)
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import _probe_gen as pg

MINCHO = pg.MINCHO
esc = pg.esc

# sweep grid: hps (half-pt) x hpsRaise (half-pt, None = omit)
HPS = [6, 8, 10, 14, 21]
RAISE = [None, 6, 12, 18, 24, 30]

def rpr(sz):
    return f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/><w:sz w:val="{sz}"/>'

def plain(txt, sz):
    r = rpr(sz)
    return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

def ruby_para(idx, hps, raise_, sz):
    r = rpr(sz)
    raise_el = f'<w:hpsRaise w:val="{raise_}"/>' if raise_ is not None else ''
    ruby = ('<w:r><w:ruby>'
            f'<w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="{hps}"/>{raise_el}'
            f'<w:hpsBaseText w:val="{sz}"/><w:lid w:val="ja-JP"/></w:rubyPr>'
            f'<w:rt><w:r><w:rPr><w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
            f'<w:sz w:val="{hps}"/></w:rPr><w:t>かんじ</w:t></w:r></w:rt>'
            f'<w:rubyBase><w:r><w:rPr>{r}</w:rPr><w:t>漢字</w:t></w:r></w:rubyBase>'
            '</w:ruby></w:r>')
    # unique numeric prefix so measure_pagination matches each para reliably
    return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">条{idx}項　</w:t></w:r>'
            + ruby +
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">　の件</w:t></w:r></w:p>')

def build(sz, name):
    paras = []
    idx = 0
    configs = []
    for hps in HPS:
        for raise_ in RAISE:
            idx += 1
            configs.append((idx, hps, raise_))
            paras.append(ruby_para(idx, hps, raise_, sz))
            paras.append(plain(f"控{idx}行目の通常段落である。", sz))
    body = "".join(paras) + pg.sectpr()
    pg.write_docx(pg.out(name), pg.doc(body), sz=str(sz))
    print(f"wrote {name}: {idx} ruby configs (sz={sz})")
    return configs

if __name__ == "__main__":
    cfg = build(21, "probervsweep21_rubygridsweep.docx")
    build(24, "probervsweep24_rubygridsweep.docx")
    import json
    json.dump(cfg, open(os.path.join(os.path.dirname(__file__), "_ruby_sweep_configs.json"), "w"))
