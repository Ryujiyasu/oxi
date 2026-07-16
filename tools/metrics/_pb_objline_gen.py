# -*- coding: utf-8 -*-
"""Pin the inline w:object line-height rule (the REPORT2 B-2/B-3 2-arm model).

Derived on forms__00042714 (+-0.4pt): mixed text+obj line H =
max(normal_line, obj_h + (normal_line - text_ascent)); solo object para
H = obj_h + raw mark descent. This sweep pins the text_ascent convention
(hhea asc / hhea asc+lineGap / win asc) and the solo-descent convention.

One docx, one render: for each config (obj_h x line rule x mixed/solo) a
sandwich [marker plain] [test para] [marker plain]; all paras before=0
after=0 direct so consecutive Info(6) y-gaps read the line heights
directly (0.75 quantized; 18 points overdetermine the conventions).
Arial 11 (hhea asc 9.958, +gap 10.318, win asc 9.958+? see fontTools) --
NOTE Arial win asc == hhea asc (1854), so the discriminating font is
CALIBRI (hhea 1536+452 gap vs win 1950): the c-series repeats the sweep
with Calibri 11 runs.

Usage:
  python _pb_objline_gen.py gen
  python _pb_objline_gen.py measure
"""
import os, sys, zipfile, json, base64

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_objline")

# 1x1 white PNG
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAE"
    "hQGAhKmMIQAAAABJRU5ErkJggg==")

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:o="urn:schemas-microsoft-com:office:office"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '<Relationship Id="rImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>'
           '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/>'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr>'
          '</w:style></w:styles>')


def rpr(font):
    return f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="22"/>'


def obj_run(h_pt, font):
    w_pt = 60.0
    return (f'<w:r><w:rPr>{rpr(font)}</w:rPr>'
            f'<w:object w:dxaOrig="{int(w_pt*20)}" w:dyaOrig="{int(h_pt*20)}">'
            f'<v:shape id="s1" type="#_x0000_t75" style="width:{w_pt}pt;height:{h_pt}pt">'
            f'<v:imagedata r:id="rImg" o:title=""/></v:shape>'
            f'</w:object></w:r>')


def para(inner, line_tw, font):
    return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" '
            f'w:line="{line_tw}" w:lineRule="auto"/><w:rPr>{rpr(font)}</w:rPr></w:pPr>'
            f'{inner}</w:p>')


def marker(tag, font):
    return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" '
            f'w:line="240" w:lineRule="auto"/><w:rPr>{rpr(font)}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr(font)}</w:rPr><w:t>MK{tag}</w:t></w:r></w:p>')


# configs: (tag, font, obj_h, line_tw, mixed)
CASES = []
for font, fkey in (("Arial", "a"), ("Calibri", "c")):
    for oh in (12, 18, 24):
        for lt in (240, 276, 360):
            for mixed in (True, False):
                tag = f"{fkey}{oh:02d}{lt}{'m' if mixed else 's'}"
                CASES.append((tag, font, float(oh), lt, mixed))


def build():
    body = []
    for i, (tag, font, oh, lt, mixed) in enumerate(CASES):
        body.append(marker(f"{i:02d}A", font))
        inner = obj_run(oh, font)
        if mixed:
            inner = (f'<w:r><w:rPr>{rpr(font)}</w:rPr><w:t xml:space="preserve">Tx </w:t></w:r>'
                     + inner)
        body.append(para(inner, lt, font))
    body.append(marker("ZZZ", "Arial"))
    b = ''.join(body) + (
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="720" w:right="1440" w:bottom="720" '
        'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{b}</w:body></w:document>')


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    with zipfile.ZipFile(os.path.join(OUTDIR, 'objline.docx'), 'w',
                         zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOCRELS)
        z.writestr('word/document.xml', build())
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/media/image1.png', PNG)
    print(f'generated objline.docx ({len(CASES)} configs) -> {os.path.abspath(OUTDIR)}')


def measure():
    import win32com.client
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    rows = []
    try:
        p = os.path.abspath(os.path.join(OUTDIR, 'objline.docx'))
        d = word.Documents.Open(p, ReadOnly=True)
        try:
            for i in range(1, d.Paragraphs.Count + 1):
                r = d.Paragraphs(i).Range
                t = ''.join(ch for ch in r.Text if ch.isprintable()).strip()
                rs = d.Range(r.Start, r.Start)
                rows.append((t[:12], rs.Information(3), round(rs.Information(6), 2)))
        finally:
            d.Close(False)
    finally:
        word.Quit()
    json.dump(rows, open(os.path.join(OUTDIR, '_paras.json'), 'w'), indent=0)
    # readout: test para height = next_marker.y - test.y (same page only)
    print(f"{'tag':>9} {'obj_h':>5} {'line':>4} {'mode':>5} {'H(pt)':>7}")
    res = {}
    for i, (tag, font, oh, lt, mixed) in enumerate(CASES):
        # rows: marker(2i), test(2i+1), next marker(2i+2)
        mi, ti, ni = 2 * i, 2 * i + 1, 2 * i + 2
        if ni >= len(rows): break
        (mt, mp, my), (tt, tp, ty), (nt, np_, ny) = rows[mi], rows[ti], rows[ni]
        h = (ny - ty) if np_ == tp else None
        res[tag] = h
        print(f"{tag:>9} {oh:5.0f} {lt:4d} {'mixed' if mixed else 'solo':>5} "
              f"{h if h is not None else 'PGBRK':>7}")
    json.dump(res, open(os.path.join(OUTDIR, '_measure.json'), 'w'), indent=1)
    print('wrote _measure.json')


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        gen()
    else:
        measure()
