"""S492 canary — true OFF (var unset) vs ON, repros + real bottom-N docs vs Word.

Renders each docx twice (env var absent vs =1) and reports L1 char counts.
Verifies: (a) jc=both UNCHANGED OFF->ON (S492 only touches non-justified);
(b) jc=left ON matches Word; (c) real docs 683f/0e7af/d77a jc=left lines match Word.
"""
import os, json, subprocess, glob, re
import win32com.client as w32

BIN = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
WD_VPOS = 6
ALIGN = {0: 'left', 1: 'center', 2: 'right', 3: 'both', 4: 'distribute'}


def render(docx, on):
    env = dict(os.environ)
    env.pop('OXI_S492_JCNATURAL', None)
    if on:
        env['OXI_S492_JCNATURAL'] = '1'
    out = 'c:/tmp/_can_%s_%d.json' % (re.sub(r'\W', '', os.path.basename(docx))[:18], on)
    subprocess.run([BIN, docx, 'c:/tmp/_can_x', '--dump-layout=' + out],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
    return out


def oxi_paras(jsonpath):
    d = json.load(open(jsonpath, encoding='utf-8'))
    paras = {}
    for pg in d['pages']:
        for e in pg['elements']:
            if e['type'] != 'text' or e.get('para_idx') is None:
                continue
            paras.setdefault(e['para_idx'], []).append(e)
    res = {}
    for pi, els in paras.items():
        y0 = sorted(set(round(e['y'], 1) for e in els))[0]
        l1 = [e for e in els if abs(e['y'] - y0) < 2]
        full = sorted(els, key=lambda e: (round(e['y'], 1), e['x']))
        pref = re.sub(r'\s', '', ''.join(e['text'] for e in full))[:12]
        res[pi] = (len(l1), pref)
    return res


# --- repros ---
DIR = 'tools/golden-test/repros/breakflip_jc'
repro_word = {'bf_comma_left': 36, 'bf_open_kak_left': 37, 'bf_close_paren_left': 36,
              'rd_d6_left': 37, 'rd_d10_left': 37, 'rd_mix_left': 37,
              'bf_comma_both': 38, 'rd_d6_both': 38, 'rd_mix_both': 38}
print("=== repros: OFF vs ON vs Word ===")
print("%-22s %4s %4s %4s  %s" % ('variant', 'OFF', 'ON', 'Word', 'note'))
for v, w in repro_word.items():
    off = oxi_paras(render(os.path.abspath(f'{DIR}/{v}.docx'), 0))
    on = oxi_paras(render(os.path.abspath(f'{DIR}/{v}.docx'), 1))
    o = off[0][0]; n = on[0][0]
    if v.endswith('_both'):
        note = 'just OFF==ON OK' if o == n else 'BAD: justified changed'
    else:
        note = 'ON==Word OK' if n == w else 'ON off by %+d (burasagari?)' % (n - w)
    print("%-22s %4d %4d %4d  %s" % (v, o, n, w, note))

# --- real docs ---
DOCS = {}
for stem in ['683ffcab86e2', '0e7af1ae8f21', 'd77a58485f16', 'b837808d0555']:
    g = glob.glob('pipeline_data/golden_per_page/%s*_p1.docx' % stem)
    if g:
        DOCS[stem] = os.path.abspath(g[0])

word = w32.DispatchEx('Word.Application'); word.Visible = False
try:
    for stem, docx in DOCS.items():
        off = oxi_paras(render(docx, 0))
        on = oxi_paras(render(docx, 1))
        off_by = {p[1]: p[0] for p in off.values()}
        on_by = {p[1]: p[0] for p in on.values()}
        doc = word.Documents.Open(docx, ReadOnly=True)
        rows = []
        try:
            for p in doc.Paragraphs:
                rng = p.Range; txt = rng.Text
                clean = txt.replace('\r', '').replace('\x07', '').replace('\n', '')
                if len(clean) < 20:
                    continue
                start, end = rng.Start, rng.End
                y0 = doc.Range(start, start).Information(WD_VPOS)
                yN = doc.Range(max(start, end - 1), max(start, end - 1)).Information(WD_VPOS)
                if (yN - y0) <= 2:
                    continue
                al = ALIGN.get(p.Alignment, str(p.Alignment))
                n = 0
                for i in range(len(txt)):
                    if txt[i] in ('\r', '\n', '\x07'):
                        continue
                    if doc.Range(start + i, start + i).Information(WD_VPOS) > y0 + 2:
                        break
                    n += 1
                pref = re.sub(r'\s', '', clean)[:12]
                rows.append((al, n, off_by.get(pref), on_by.get(pref), clean[:14]))
        finally:
            doc.Close(False)
        print("\n=== %s ===  align  Word  OFF  ON  text" % stem)
        for al, w, o, nn, t in rows:
            if o is None:
                continue
            tag = ''
            if al != 'both':
                tag = '  FIXED' if (o != w and nn == w) else ('  still off %+d' % (nn - w) if nn != w else '  was-ok')
            else:
                tag = '  just-unchanged' if o == nn else '  BAD just changed'
            print("   %-6s %4d %4d %4d  %s%s" % (al, w, o, nn, t, tag))
finally:
    word.Quit()
