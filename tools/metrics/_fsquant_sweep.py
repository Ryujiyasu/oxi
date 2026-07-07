# -*- coding: utf-8 -*-
"""Controlled sweep: fs x CONTEXT grid quantization (the probeqsizes family,
bd90b00's Kousei-Roudou-Daijin-dono specimen).

Question: a LARGE-font line inside a typed docGrid — does its consumed
height depend on CONTEXT (single line between default-size lines vs a
steady run)? bd90b00: a single 14pt line = Word ~24.95 vs Oxi 2 cells
33.0 (@pitch 16.5), while the gridquant STEADY-run sweep measured 14pt =
2 cells.

Convention-free design: each config is a sandwich
    D D D | L x k | D D D        (D = default 10.5pt auto, L = fs F auto)
and we measure ink(D3.line) -> ink(D4.line) = G(k). Both anchors are the
SAME class so per-class ink conventions cancel. G(1) = the single-line
slot total (junctions included); (G(3)-G(1))/2 = the steady pitch.
Word truth = PDF ink (fitz); Oxi = --dump-layout el.y with the same
same-class-difference computation.

Run: python tools/metrics/_fsquant_sweep.py
"""
import os, sys, json, subprocess
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

esc = pg.esc
MINCHO = pg.MINCHO
OUTDIR = os.path.join(os.environ.get("TEMP", "."), 'gridquant')
os.makedirs(OUTDIR, exist_ok=True)

def rpr(sz):
    return (f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
            f'<w:sz w:val="{sz}"/>')

def para(txt, sz):
    r = rpr(sz)
    return (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

def pagebreak():
    r = rpr("21")
    return (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:br w:type="page"/></w:r></w:p>')

FS = ["24", "28", "32", "42", "56"]   # 12 / 14 / 16 / 21 / 28 pt
KS = [1, 3]
PITCHES = [330, 360]

def build(pitch):
    body = []
    configs = []
    first = True
    for szv in FS:
        for k in KS:
            if not first:
                body.append(pagebreak())
            first = False
            cid = f'{szv}k{k}'
            configs.append(cid)
            tag = f'ｘ{len(configs):02d}'
            body.append(para(f'Ｄ１{tag}', "21"))
            body.append(para(f'Ｄ２{tag}', "21"))
            body.append(para(f'Ｄ３{tag}', "21"))
            for j in range(k):
                body.append(para(f'Ｌ{j}{tag}', szv))
            body.append(para(f'Ｄ４{tag}', "21"))
            body.append(para(f'Ｄ５{tag}', "21"))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="851" w:left="1134" '
            'w:header="851" w:footer="567" w:gutter="0"/>'
            f'<w:docGrid w:type="lines" w:linePitch="{pitch}"/></w:sectPr>')
    body.append(sect)
    docx = os.path.join(OUTDIR, f'fsquant_{pitch}.docx')
    pg.write_docx(docx, pg.doc(''.join(body)))
    return docx, configs

def word_measure(docx, pdf):
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(docx, ReadOnly=True, AddToRecentFiles=False)
    doc.ExportAsFixedFormat(pdf, 17)
    doc.Close(False)
    word.Quit()

def pdf_anchors(pdf):
    import fitz
    d = fitz.open(pdf)
    out = {}
    for p in range(len(d)):
        for b in d[p].get_text('dict')['blocks']:
            for l in b.get('lines', []):
                txt = ''.join(s['text'] for s in l.get('spans', [])).replace(' ', '')
                if txt.startswith(('Ｄ３', 'Ｄ４')):
                    out[txt[:5]] = round(l['bbox'][1], 2)
    return out

def oxi_anchors(docx):
    dumpjson = os.path.join(OUTDIR, 'fsquant_oxi.json')
    env = dict(os.environ)
    for v in ('OXI_ROWBOX2', 'OXI_OFFSLOT_INK', 'OXI_S638_ALL'):
        env[v] = '1'
    r = subprocess.run(['./tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe',
                        docx, os.path.join(OUTDIR, 'fsq_png'),
                        '--dump-layout=' + dumpjson],
                       capture_output=True, env=env)
    dump = json.load(open(dumpjson, encoding='utf-8'))
    out = {}
    for pi, page in enumerate(dump['pages']):
        rows = {}
        for e in page['elements']:
            if e['type'] == 'text' and e.get('text', '').strip():
                rows.setdefault(round(e['y'], 2), []).append((e['x'], e['text']))
        for y, frs in rows.items():
            t = ''.join(t for _, t in sorted(frs)).replace('　', '')
            if t.startswith(('Ｄ３', 'Ｄ４')):
                out[t[:5]] = y
    return out

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    for pitch in PITCHES:
        docx, configs = build(pitch)
        pdf = docx.replace('.docx', '.pdf')
        word_measure(docx, pdf)
        wa = pdf_anchors(pdf)
        oa = oxi_anchors(docx)
        ppt = pitch / 20.0
        print(f'\n=== pitch {pitch} ({ppt}pt) ===')
        print(f'{"fs":>4} {"k":>2} {"G_word":>8} {"G_oxi":>8} {"d":>7}   cells_w   cells_o')
        g = {}
        for i, cid in enumerate(configs):
            tag = f'ｘ{i+1:02d}'
            try:
                gw = round(wa[f'Ｄ４{tag}'] - wa[f'Ｄ３{tag}'], 2)
                go = round(oa[f'Ｄ４{tag}'] - oa[f'Ｄ３{tag}'], 2)
            except KeyError:
                print(f'{cid}: MISSING anchors')
                continue
            g[cid] = (gw, go)
            fs = int(cid.split('k')[0]) / 2
            k = int(cid.split('k')[1])
            print(f'{fs:4.0f} {k:2d} {gw:8.2f} {go:8.2f} {go-gw:+7.2f}   {gw/ppt:7.3f}   {go/ppt:7.3f}')
        print('--- derived (steady = (G3-G1)/2; single slot = G1 - D_pitch):')
        for szv in FS:
            k1 = g.get(f'{szv}k1'); k3 = g.get(f'{szv}k3')
            if not (k1 and k3):
                continue
            fs = int(szv) / 2
            sw = (k3[0] - k1[0]) / 2
            so = (k3[1] - k1[1]) / 2
            single_w = k1[0] - ppt   # G1 = D3->L + L->D4 = L-slot + 1 D-line
            single_o = k1[1] - ppt
            print(f'  fs={fs:4.1f}: steady W={sw:6.2f} O={so:6.2f} | single-slot W={single_w:6.2f} O={single_o:6.2f}')
