"""S492/S476 — b837 p7 +18pt target (S452c's single cleanest per-page-step lever).
Oxi fits one fewer char/line in the 本府は paragraph (p6) -> +1 line -> spills to p7
-> all of p7 shifts +18pt. Measure Word per-line char counts vs Oxi for that para to
see if Oxi UNDER-packs (capacity-K too low) and whether it's char-width or break-policy.
No code change.
"""
import os, glob, subprocess, json, re
import win32com.client as w32

BIN = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
WD_VPOS = 6
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/b837*.docx')[0])
KEY = '本府'  # paragraph start marker (S452c: 「本府は…」)


def oxi_para_lines():
    out = 'c:/tmp/_b837.json'
    subprocess.run([BIN, DOCX, 'c:/tmp/_b837_x', '--dump-layout=' + out],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    d = json.load(open(out, encoding='utf-8'))
    paras = {}
    for pg in d['pages']:
        for e in pg['elements']:
            if e['type'] != 'text' or e.get('para_idx') is None:
                continue
            paras.setdefault(e['para_idx'], []).append(e)
    # find para whose text contains KEY near start
    for pi, els in paras.items():
        full = sorted(els, key=lambda e: (round(e['y'], 1), e['x']))
        txt = ''.join(e['text'] for e in full)
        if txt.startswith(KEY) or txt[:6].find(KEY) >= 0:
            # group into lines by y
            lines = {}
            for e in full:
                lines.setdefault(round(e['y'], 1), []).append(e)
            line_counts = [len(''.join(x['text'] for x in lines[y])) for y in sorted(lines)]
            return txt, line_counts
    return None, None


otxt, ocounts = oxi_para_lines()
print("Oxi 本府は para: %d lines, counts=%s" % (len(ocounts) if ocounts else 0, ocounts))
print("Oxi text[:40]:", repr(otxt[:40]) if otxt else None)

word = w32.DispatchEx('Word.Application'); word.Visible = False
try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        for p in doc.Paragraphs:
            rng = p.Range; txt = rng.Text
            clean = txt.replace('\r', '').replace('\x07', '').replace('\n', '')
            if not clean.startswith(KEY):
                continue
            start = rng.Start
            y0 = doc.Range(start, start).Information(WD_VPOS)
            # per-line counts: walk chars, new line when Y jumps
            counts = []
            cur = 0
            prev_y = y0
            for i in range(len(txt)):
                ch = txt[i]
                if ch in ('\r', '\n', '\x07'):
                    continue
                y = doc.Range(start + i, start + i).Information(WD_VPOS)
                if y > prev_y + 2:
                    counts.append(cur); cur = 0; prev_y = y
                cur += 1
            if cur:
                counts.append(cur)
            print("\nWord 本府は para: %d lines, counts=%s" % (len(counts), counts))
            print("Word text[:40]:", repr(clean[:40]))
            print("\n=== verdict ===")
            if ocounts:
                print("Word lines=%d vs Oxi lines=%d (Oxi %+d lines)" % (len(counts), len(ocounts), len(ocounts) - len(counts)))
                for i in range(max(len(counts), len(ocounts))):
                    w = counts[i] if i < len(counts) else '-'
                    o = ocounts[i] if i < len(ocounts) else '-'
                    print("  line %d: Word=%s Oxi=%s" % (i + 1, w, o))
            break
        else:
            print("Word: 本府は para not found")
    finally:
        doc.Close(False)
finally:
    word.Quit()
