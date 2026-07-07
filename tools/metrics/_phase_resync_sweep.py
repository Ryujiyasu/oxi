# -*- coding: utf-8 -*-
"""Grid-phase re-sync sweep: does an AUTO line following an OFF-PHASE exact
leader consume relative-ceil cells (Oxi model) or re-sync to ABSOLUTE page
grid slots in Word? Simultaneously tests the exact->auto boundary (leader
330 = on-phase control).

bd90b00 context: [exact-370 titles (+2.0 phase)] -> [auto 年月日 (1-cell)]
-> [auto 大臣殿 14pt (2-cell)]; real doc shows the pair = W 48.0 / O 49.5.

Per config: [D1 D2 D3(auto 10.5)] [leader exact-L] [T auto fs F] [D4 D5]
G(D3->D4 ink) = 16.5 + L/20 + h(T). Same-class anchors -> conventions
cancel.

Run: python tools/metrics/_phase_resync_sweep.py
"""
import os, sys, json, subprocess
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.environ.get("TEMP", "."), 'gridquant')
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, 'phaseresync.docx')
PDF = os.path.join(OUTDIR, 'phaseresync.pdf')
MINCHO = pg.MINCHO

def rpr(sz):
    return (f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
            f'<w:sz w:val="{sz}"/>')

def para(txt, sz, spacing='<w:spacing w:line="240" w:lineRule="auto"/>'):
    r = rpr(sz)
    return (f'<w:p><w:pPr>{spacing}<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{pg.esc(txt)}</w:t></w:r></w:p>')

def pagebreak():
    return ('<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:br w:type="page"/></w:r></w:p>')

LEADERS = [330, 370, 410, 290]     # 16.5(+0) 18.5(+2) 20.5(+4) 14.5(-2)
TFS = ["21", "28"]                 # 10.5 (1-cell) / 14 (2-cell)

CONFIGS = [(ld, fs) for ld in LEADERS for fs in TFS]

def build():
    body = []
    for i, (ld, fs) in enumerate(CONFIGS):
        if i:
            body.append(pagebreak())
        tag = f'ｘ{i+1:02d}'
        body.append(para(f'Ｄ１{tag}', "21"))
        body.append(para(f'Ｄ２{tag}', "21"))
        body.append(para(f'Ｄ３{tag}', "21"))
        body.append(para(f'ＬＤ{tag}', "21",
                         f'<w:spacing w:line="{ld}" w:lineRule="exact"/>'))
        body.append(para(f'Ｔ挿{tag}', fs))
        body.append(para(f'Ｄ４{tag}', "21"))
        body.append(para(f'Ｄ５{tag}', "21"))
    body.append('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1134" w:right="1134" w:bottom="851" w:left="1134" '
                'w:header="851" w:footer="567" w:gutter="0"/>'
                '<w:docGrid w:type="lines" w:linePitch="330"/></w:sectPr>')
    pg.write_docx(DOCX, pg.doc(''.join(body)))

def anchors_pdf():
    import fitz
    d = fitz.open(PDF)
    out = {}
    for p in range(len(d)):
        for b in d[p].get_text('dict')['blocks']:
            for l in b.get('lines', []):
                txt = ''.join(s['text'] for s in l.get('spans', [])).replace(' ', '')
                if txt.startswith(('Ｄ３', 'Ｄ４')):
                    out[txt[:5]] = round(l['bbox'][1], 2)
    return out

def anchors_oxi():
    dumpjson = os.path.join(OUTDIR, 'phaseresync_oxi.json')
    env = dict(os.environ)
    for v in ('OXI_ROWBOX2', 'OXI_OFFSLOT_INK', 'OXI_S638_ALL'):
        env[v] = '1'
    subprocess.run(['./tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe',
                    DOCX, os.path.join(OUTDIR, 'phaseresync_png'),
                    '--dump-layout=' + dumpjson], capture_output=True, env=env)
    dump = json.load(open(dumpjson, encoding='utf-8'))
    out = {}
    for page in dump['pages']:
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
    build()
    print('wrote', DOCX)
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(DOCX, ReadOnly=True, AddToRecentFiles=False)
    doc.ExportAsFixedFormat(PDF, 17)
    doc.Close(False)
    word.Quit()
    wa = anchors_pdf()
    oa = anchors_oxi()
    print(f'{"leader":>7} {"T_fs":>5} {"G_word":>8} {"G_oxi":>8} {"d":>7}  hT_w    hT_o   (hT = G-16.5-leader)')
    for i, (ld, fs) in enumerate(CONFIGS):
        tag = f'ｘ{i+1:02d}'
        try:
            gw = round(wa[f'Ｄ４{tag}'] - wa[f'Ｄ３{tag}'], 2)
            go = round(oa[f'Ｄ４{tag}'] - oa[f'Ｄ３{tag}'], 2)
            base = 16.5 + ld / 20.0
            print(f'{ld:7d} {int(fs)/2:5.1f} {gw:8.2f} {go:8.2f} {go-gw:+7.2f}  {gw-base:6.2f}  {go-base:6.2f}')
        except KeyError:
            print(f'{ld:7d} {int(fs)/2:5.1f} MISSING')
