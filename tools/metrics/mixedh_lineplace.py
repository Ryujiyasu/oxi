# -*- coding: utf-8 -*-
"""mixed-height line placement precision harness (Ra, 2026-06-24).

Compares Word (COM, per-paragraph Y via Information(6) with the R30 collapsed-
start fix) against Oxi (--dump-layout per-element Y) for the three mixed-height
contexts the precision pass targets:

  - math   : a display equation between two body paragraphs. The equation
             paragraph's reserved height = Y(below) - Y(equation_top).
  - empty  : an empty paragraph between two body paragraphs (grid + no-grid).
  - heading: a large heading paragraph followed by body (cross-paragraph gap).

For each repro it reports, per measured boundary, the Word gap, the Oxi gap and
the signed delta (Oxi - Word). cp932-safe stdout.

Usage:
  python tools/metrics/mixedh_lineplace.py math      # build+measure math repros
  python tools/metrics/mixedh_lineplace.py <doc.docx>  # measure a single doc
"""
import os, sys, io, json, zipfile, subprocess, tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
OUT = os.path.join('c:/tmp', 'mixedh')
os.makedirs(OUT, exist_ok=True)

WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
MNS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')


def r(t):
    return '<m:r><m:t>%s</m:t></m:r>' % t


# OMML structures of varying intrinsic height (reused from S526 + extras).
MATH = {
    'x':      r('x'),                                                            # 1 line tall
    'sup':    '<m:sSup><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup>' % (r('x'), r('2')),
    'frac':   '<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>' % (r('a'), r('b')),
    'nest':   '<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>' % (
              '<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>' % (r('a'), r('b')), r('c')),  # 3 levels
    'rad':    '<m:rad><m:deg/><m:e>%s</m:e></m:rad>' % r('x'),
    'sum':    '<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>%s</m:sub><m:sup>%s</m:sup><m:e>%s</m:e></m:nary>' % (
              r('1'), r('n'), r('i')),                                          # tall (n-ary)
    'matrix': '<m:m><m:mr><m:e>%s</m:e><m:e>%s</m:e></m:mr><m:mr><m:e>%s</m:e><m:e>%s</m:e></m:mr></m:m>' % (
              r('a'), r('b'), r('c'), r('d')),                                  # 2 rows tall
}

BODY = ('<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        '<w:sz w:val="21"/></w:rPr></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        '<w:sz w:val="21"/></w:rPr><w:t>%s</w:t></w:r></w:p>')


def build_math(name, omml):
    eqpara = '<w:p><m:oMathPara %s><m:oMath>%s</m:oMath></m:oMathPara></w:p>' % (MNS, omml)
    body = (BODY % 'TOP') + eqpara + (BODY % 'BOTTOM')
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s %s><w:body>%s%s</w:body></w:document>' % (WNS, MNS, body, sect))
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', doc)
    return p


def build_generic(name, body_xml, grid_pitch=None):
    grid = ('<w:docGrid w:type="lines" w:linePitch="%d"/>' % grid_pitch) if grid_pitch else ''
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/>%s</w:sectPr>' % grid)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s %s><w:body>%s%s</w:body></w:document>' % (WNS, MNS, body_xml, sect))
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', doc)
    return p


def para(text, font='ＭＳ 明朝', sz=21, bold=False, empty=False):
    b = '<w:b/>' if bold else ''
    rpr = '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/>%s' % (font, font, font, sz, b)
    if empty:
        return '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr></w:p>' % rpr
    return '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr><w:t>%s</w:t></w:r></w:p>' % (rpr, rpr, text)


def measure_seq(name, body_xml, grid_pitch, labels):
    """Generic: measure Word & Oxi Y of each labelled paragraph; report gaps."""
    dx = build_generic(name + '.docx', body_xml, grid_pitch)
    wp = word_paras(dx)
    els = oxi_elements(oxi_dump(dx))
    wy, oy = {}, {}
    for lab in labels:
        wy[lab] = next((p['y'] for p in wp if p['text'].startswith(lab)), None)
        oy[lab] = next((e['y'] for e in els if e['text'].startswith(lab)), None)
    g = 'grid%d' % grid_pitch if grid_pitch else 'nogrid'
    parts = []
    for i in range(1, len(labels)):
        a, b = labels[i - 1], labels[i]
        wg = round(wy[b] - wy[a], 2) if (wy[a] is not None and wy[b] is not None) else None
        og = round(oy[b] - oy[a], 2) if (oy[a] is not None and oy[b] is not None) else None
        d = round(og - wg, 2) if (wg is not None and og is not None) else None
        parts.append('%s->%s W=%-6s O=%-6s d=%+-6.2f' % (a, b, wg, og, d if d is not None else 0.0))
    print('%-22s %-8s | %s' % (name, g, '  '.join(parts)))


def measure_empty():
    print('=== EMPTY: empty-paragraph reserved height between body lines ===')
    # body(11pt) / empty(varying sz) / body(11pt), grid + nogrid
    for esz, elabel in [(22, 'e11'), (28, 'e14'), (48, 'e24')]:
        body = para('AAA', sz=22) + para('', sz=esz, empty=True) + para('BBB', sz=22)
        for gp in (None, 360):
            measure_seq('empty_%s' % elabel, body, gp, ['AAA', 'BBB'])


def measure_heading():
    print('=== HEADING: heading->body transition gap ===')
    for hsz, hlabel in [(28, 'h14'), (40, 'h20'), (60, 'h30')]:
        body = para('TOP', sz=22) + para('HEAD', sz=hsz, bold=True) + para('BODY', sz=22)
        for gp in (None, 360):
            measure_seq('head_%s' % hlabel, body, gp, ['TOP', 'HEAD', 'BODY'])


def word_paras(docx):
    """Per-paragraph (text, page, y) using collapsed-start Information (R30)."""
    import win32com.client, pythoncom
    pythoncom.CoInitialize()
    w = win32com.client.DispatchEx('Word.Application')
    w.Visible = False
    out = []
    try:
        d = w.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        for p in d.Paragraphs:
            rng = p.Range
            start = d.Range(rng.Start, rng.Start)
            page = start.Information(3)
            y = float(start.Information(6))
            txt = rng.Text.strip().replace('\r', '')
            out.append({'text': txt[:24], 'page': int(page), 'y': round(y, 2)})
        d.Close(False)
    finally:
        w.Quit()
    return out


def oxi_dump(docx):
    with tempfile.TemporaryDirectory(prefix='mh_') as tmp:
        dp = os.path.join(tmp, 'l.json')
        op = os.path.join(tmp, 'o')
        subprocess.run([EXE, os.path.abspath(docx), op, '--dump-layout=' + dp],
                       capture_output=True, text=True)
        with open(dp, encoding='utf-8') as f:
            return json.load(f)


def oxi_elements(dump):
    """Flatten to (page, y, h, text, type, para_idx) in document order."""
    els = []
    for pg in dump.get('pages', []):
        for e in pg['elements']:
            els.append({'page': pg['page'], 'y': round(e['y'], 2), 'h': round(e['h'], 2),
                        'text': (e.get('text') or '')[:24], 'type': e['type'],
                        'para_idx': e.get('para_idx')})
    return els


def word_math_ink(docx):
    """Word's true math ink vertical extent (pt) from the exported PDF page 1.
    The equation is the only content between TOP and BOTTOM text rows; we take
    the ink band that is neither TOP's nor BOTTOM's row."""
    import win32com.client, pythoncom, fitz
    pdf = os.path.splitext(docx)[0] + '.pdf'
    pythoncom.CoInitialize()
    w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    doc = fitz.open(pdf); pg = doc[0]
    spans = []
    for b in pg.get_text('dict')['blocks']:
        for ln in b.get('lines', []):
            for sp in ln['spans']:
                spans.append((sp['bbox'][1], sp['bbox'][3], sp['text']))
    doc.close()
    tops = [s for s in spans if 'TOP' in s[2]]
    bots = [s for s in spans if 'BOTTOM' in s[2]]
    if not tops or not bots:
        return None
    t_bot = max(s[1] for s in tops); b_top = min(s[0] for s in bots)
    mid = [s for s in spans if s[0] >= t_bot - 1 and s[1] <= b_top + 1
           and 'TOP' not in s[2] and 'BOTTOM' not in s[2]]
    if not mid:
        return None
    return round(max(s[1] for s in mid) - min(s[0] for s in mid), 2)


def measure_math():
    print('=== MATH: display-equation reserved height (Word vs Oxi) ===')
    print('reserve = Y(BOTTOM)-Y(eq_top);  ink = rendered math vertical extent')
    print('%-8s | %-7s %-7s %-7s | %-7s %-7s %-7s | %-7s %-7s'
          % ('name', 'W_resv', 'W_ink', 'W_lead', 'O_resv', 'O_ink', 'O_lead', 'd_resv', 'd_ink'))
    for name, omml in MATH.items():
        dx = build_math('mh_' + name + '.docx', omml)
        wp = word_paras(dx)
        wy = {p['text']: p['y'] for p in wp}
        ytop = wy.get('TOP'); ybot = wy.get('BOTTOM')
        yeq = next((p['y'] for p in wp if p['text'] not in ('TOP', 'BOTTOM')), ytop)
        w_resv = round(ybot - yeq, 2)
        w_ink = word_math_ink(dx)
        w_lead = round(w_resv - w_ink, 2) if w_ink else None
        els = oxi_elements(oxi_dump(dx))
        oy_top = next((e['y'] for e in els if e['text'] == 'TOP'), None)
        oy_bot = next((e['y'] for e in els if e['text'] == 'BOTTOM'), None)
        mid = [e for e in els if e['text'] not in ('TOP', 'BOTTOM')]
        oy_eqtop = min(e['y'] for e in mid) if mid else oy_top
        o_ink = round(max(e['y'] + e['h'] for e in mid) - min(e['y'] for e in mid), 2) if mid else None
        o_resv = round(oy_bot - oy_eqtop, 2) if oy_bot is not None else None
        o_lead = round(o_resv - o_ink, 2) if (o_resv and o_ink) else None
        d_resv = round(o_resv - w_resv, 2) if o_resv is not None else None
        d_ink = round(o_ink - w_ink, 2) if (o_ink and w_ink) else None
        print('%-8s | %-7s %-7s %-7s | %-7s %-7s %-7s | %+-7.2f %s'
              % (name, w_resv, w_ink, w_lead, o_resv, o_ink, o_lead,
                 d_resv if d_resv is not None else 0.0,
                 ('%+.2f' % d_ink) if d_ink is not None else 'n/a'))


def _ink_bands(png, dpi):
    import numpy as np
    from PIL import Image
    im = np.asarray(Image.open(png).convert('L'), dtype=np.float32)
    rows = (im < 128).sum(axis=1) > 0
    bands = []
    i, n = 0, len(rows)
    while i < n:
        if rows[i]:
            j = i
            while j < n and (rows[j] or (j + 1 < n and rows[j + 1])):
                j += 1
            bands.append((i, j - 1)); i = j
        else:
            i += 1
    return [(t / dpi * 72.0, b / dpi * 72.0) for t, b in bands]


def _word_render_png(docx, dpi):
    import win32com.client, pythoncom, fitz
    pdf = os.path.splitext(docx)[0] + '.pdf'
    png = os.path.splitext(docx)[0] + '_wp.png'
    if os.path.exists(pdf):
        try: os.remove(pdf)
        except OSError: pass
    pythoncom.CoInitialize()
    w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    doc = fitz.open(pdf)
    doc[0].get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72)).save(png)
    doc.close()
    return png


def measure_mathpix():
    """AUTHORITATIVE display-equation reserve check via pixel ink bands (Word
    PDF vs Oxi PNG). Info(6) is unreliable for the equation row (CLAUDE.md
    caveat) — this measures TOP->BOTTOM ink distance, the true reserved space.
    Validates the S652 baseline-derived advance."""
    DPI = 300
    print('=== MATH (pixel): TOP->BOTTOM reserved space, Word vs Oxi ===')
    print('%-8s | %-7s %-7s %-8s | %-7s %-8s' % ('name', 'Word', 'OXI', 'd', 'OXI_old', 'd_old'))
    emax = 0.0
    for name, omml in MATH.items():
        dx = build_math('mhp_' + name + '.docx', omml)
        op = os.path.join(OUT, 'mhp_' + name)
        env = dict(os.environ)
        subprocess.run([EXE, os.path.abspath(dx), op, str(DPI)], capture_output=True, text=True, env=env)
        ob = _ink_bands(op + '_p1.png', DPI)
        env_old = dict(os.environ); env_old['OXI_S652_DISABLE'] = '1'
        subprocess.run([EXE, os.path.abspath(dx), op + '_old', str(DPI)], capture_output=True, text=True, env=env_old)
        ob_old = _ink_bands(op + '_old_p1.png', DPI)
        wb = _ink_bands(_word_render_png(dx, DPI), DPI)
        def t2b(b):
            return round(b[-1][0] - b[0][0], 2) if len(b) >= 3 else None
        w, o, oo = t2b(wb), t2b(ob), t2b(ob_old)
        d = round(o - w, 2) if (w and o) else None
        do = round(oo - w, 2) if (w and oo) else None
        if d is not None:
            emax = max(emax, abs(d))
        print('%-8s | %-7s %-7s %+-8.2f | %-7s %+-8.2f'
              % (name, w, o, d if d is not None else 0.0, oo, do if do is not None else 0.0))
    print('--- S652 max|d| = %.2f pt (was up to +16.08) ---' % emax)


def measure_doc(docx):
    print('=== %s ===' % docx)
    wp = word_paras(docx)
    els = oxi_elements(oxi_dump(docx))
    print('-- Word paragraphs --')
    for i, p in enumerate(wp):
        print('  p%-3d pg%d y=%-8.2f %r' % (i, p['page'], p['y'], p['text']))
    print('-- Oxi elements (first 40) --')
    for e in els[:40]:
        print('  pg%d y=%-8.2f h=%-6.2f %-6s %r' % (e['page'], e['y'], e['h'], e['type'], e['text']))


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')
    arg = sys.argv[1] if len(sys.argv) > 1 else 'math'
    if arg == 'math':
        measure_math()
    elif arg == 'mathpix':
        measure_mathpix()
    elif arg == 'empty':
        measure_empty()
    elif arg == 'heading':
        measure_heading()
    elif arg == 'all':
        measure_math(); measure_empty(); measure_heading()
    else:
        measure_doc(arg)
