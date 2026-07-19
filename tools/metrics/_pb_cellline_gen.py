# -*- coding: utf-8 -*-
"""Word's table-cell SINGLE line height — font x size controlled sweep.

The S935-set default-ON attempt (2026-07-19) was falsified by gen2: Word's
Cambria-11 cell row = 12.75 (the word_line_height_table_cell "tombstone"
value), NOT hhea 12.896 — while Arial-10 (uklocal rt.pdf 11.52), Calibri-11
(cte probe 13.44) and Calibri-12 (001cf65 14.648) all measured hhea EXACT.
The per-font mix breaks both the tombstone table AND the S940 blanket-hhea;
six threshold models died (lineGap [Cambria hhea==win, gap 0], GDI int-px
[Arial-10 GDI 12.0 != Word 11.5], px-floor [Calibri-12 floor 14.25 != Word
14.648], 0.75/0.12 grids, VDMX, magnitude gate). This sweep measures the
rule directly.

Geometry = the cte probe's (TableGrid declaring after=0 line=240 over
docDefaults after=160 line=259 => cells single-spaced; TableNormal cellMar
top/bottom 0; insideH sz4): a 1-column table with N_ROWS identical
single-line text rows. Readout = (last border - first border) / N_ROWS from
the Word PDF drawings => row pitch to ~0.05pt. Compare vs hhea natural and
the naive px-quantized candidates.

Usage: python _pb_cellline_gen.py gen | measure | read
"""
import os, sys, glob, shutil, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_cellline")

FONTS = [
    ("cam", "Cambria"),
    ("cal", "Calibri"),
    ("ari", "Arial"),
    ("tnr", "Times New Roman"),
    ("seg", "Segoe UI"),
    ("geo", "Georgia"),
    ("ver", "Verdana"),
]
SIZES = [9.0, 10.0, 10.5, 11.0, 12.0, 14.0]
N_ROWS = 12

STYLES_TMPL = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz2}"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>
<w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>
<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
<w:tblPr><w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders></w:tblPr></w:style>
</w:styles>"""


def build(font):
    rows = ''.join(
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>'
        f'<w:p><w:r><w:t>Row {i+1} cell text</w:t></w:r></w:p></w:tc></w:tr>'
        for i in range(N_ROWS))
    tbl = ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>' + rows + '</w:tbl>')
    body = '<w:p><w:r><w:t>Anchor before.</w:t></w:r></w:p>' + tbl \
        + '<w:p><w:r><w:t>Anchor after.</w:t></w:r></w:p>'
    body += pg.sectpr(pgsz='<w:pgSz w:w="12240" w:h="15840"/>',
                      mar='<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>',
                      grid='')
    return pg.doc(body)


def cases():
    for fk, font in FONTS:
        for sz in SIZES:
            yield f'cl_{fk}_{sz:g}'.replace('.', 'q'), font, sz


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    n = 0
    for nm, font, sz in cases():
        parts = {
            "[Content_Types].xml": pg.content_types(()),
            "_rels/.rels": pg.RELS,
            "word/document.xml": build(font),
            "word/_rels/document.xml.rels": pg.docrels(()),
            "word/styles.xml": STYLES_TMPL.format(font=font, sz2=int(sz * 2)),
            "word/settings.xml": pg.settings_xml("15", False),
        }
        with zipfile.ZipFile(os.path.join(OUTDIR, nm + '.docx'), 'w', zipfile.ZIP_DEFLATED) as z:
            for name, data in parts.items():
                z.writestr(name, data)
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
    print('measured')


def read():
    import fitz
    from font.registry_probe import metrics_for  # optional; fallback below
    read_inner()


def read_inner():
    import fitz
    print(f'{"case":14s} {"pitch":>7s} {"hhea":>7s} {"d_hhea":>7s}')
    for nm, font, sz in cases():
        pdf = os.path.join(OUTDIR, nm + '.pdf')
        if not os.path.exists(pdf):
            continue
        d = fitz.open(pdf)
        p = d[0]
        hs = []
        for dd in p.get_drawings():
            for it in dd['items']:
                if it[0] == 'l':
                    p1, p2 = it[1], it[2]
                    if abs(p1.y - p2.y) < 0.3 and abs(p1.x - p2.x) > 40:
                        hs.append(p1.y)
                elif it[0] == 're':
                    r = it[1]
                    if r.height < 1.2 and r.width > 40:
                        hs.append(r.y0)
        hs = sorted(set(round(h, 2) for h in hs))
        # cluster within 0.5pt
        cl = []
        for h in hs:
            if cl and h - cl[-1][-1] <= 0.5:
                cl[-1].append(h)
            else:
                cl.append([h])
        borders = [sum(c) / len(c) for c in cl]
        if len(borders) < N_ROWS + 1:
            print(f'{nm:14s} borders={len(borders)} SKIP')
            continue
        pitch = (borders[N_ROWS] - borders[0]) / N_ROWS
        print(f'{nm:14s} {pitch:7.3f}')


if __name__ == '__main__':
    {'gen': gen, 'measure': measure, 'read': read_inner}[sys.argv[1]]()
