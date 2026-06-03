# -*- coding: utf-8 -*-
"""S492w — COM-measure Word's cell line pitch on the FLAGGED repros (adjustLineHeightInTable
=ON, docGrid linesAndChars 350=17.5pt). Determines Word's RULE: does default-snapToGrid
snap to 17.5 (grid), 15.0 (sub-grid), or 14.0 (natural)? And does snapToGrid=0 give natural?
Also dumps Oxi's pitch for the same. Writes ASCII results. cp932-safe."""
import os, glob, json, subprocess
from pathlib import Path
import win32com.client as win32

REPRO = r'c:\tmp\cellflag'
GDI = str(Path('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe').resolve())
wdVertPos = 6

man = {}
for line in open(os.path.join(REPRO, 'manifest.txt'), encoding='utf-8'):
    name, sz, s0 = line.rstrip('\n').split('\t')
    man[name] = (float(sz), s0 == '1')


def intra_pitch_from_ys(ys):
    levels = sorted(set(round(y, 1) for y in ys))
    merged = []
    for y in levels:
        if merged and y - merged[-1] < 3:
            continue
        merged.append(y)
    deltas = [round(merged[k] - merged[k - 1], 2) for k in range(1, len(merged))]
    sd = sorted(deltas)
    return (sd[len(sd) // 2] if sd else None), len(merged), sorted(set(deltas))


# --- Word COM ---
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
wres = {}
try:
    for path in sorted(glob.glob(os.path.join(REPRO, '*.docx'))):
        name = os.path.basename(path)
        doc = word.Documents.Open(path, ReadOnly=True)
        try:
            cell = doc.Tables(1).Cell(1, 1)
            rng = cell.Range
            start, end = rng.Start, rng.End - 1
            ys = []
            i = start
            while i < end:
                cr = doc.Range(i, i)
                try:
                    ys.append(float(cr.Information(wdVertPos)))
                except Exception:
                    pass
                i += 1
            wres[name] = intra_pitch_from_ys(ys)
        finally:
            doc.Close(False)
finally:
    word.Quit()


# --- Oxi GDI dump ---
def oxi_pitch(path):
    dp = 'c:/tmp/_cf_dump.json'
    subprocess.run([GDI, path, 'c:/tmp/_cf_out', '150', '--dump-layout=' + dp],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if not os.path.exists(dp):
        return (None, 0, [])
    d = json.load(open(dp, encoding='utf-8'))
    els = [e for e in d['pages'][0]['elements'] if e.get('type') == 'text']
    return intra_pitch_from_ys([e['y'] for e in els])


print("FLAGGED repro (adjustLineHeightInTable=ON, docGrid linesAndChars linePitch=350=17.5pt):")
print("natural(no-flag, S492v): 9->12, 10->13, 10.5->14, 11->14, 12->16 ; grid=17.5")
print("size  snapToGrid  Word_pitch  Oxi_pitch   Word_deltas")
for path in sorted(glob.glob(os.path.join(REPRO, '*.docx'))):
    name = os.path.basename(path)
    sz, s0 = man[name]
    wp, wn, wd = wres.get(name, (None, 0, []))
    op, on, od = oxi_pitch(path)
    print("  %4.1f  %-10s  %s      %s      %s" % (
        sz, 'snap=0' if s0 else 'default',
        ('%.2f' % wp) if wp else ' - ',
        ('%.2f' % op) if op else ' - ', wd))
print("\nINTERPRETATION: default-snapToGrid Word pitch == 17.5 => Word snaps to grid (Oxi correct);")
print("  == ~15.0 => sub-grid rule; == ~14.0 => natural (flag doesn't snap, Oxi over-snaps).")
