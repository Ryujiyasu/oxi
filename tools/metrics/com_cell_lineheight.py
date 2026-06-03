# -*- coding: utf-8 -*-
"""S492v — COM-measure intra-cell line pitch for the gen_cell_lineheight_repro docx.
For each file, open in Word, take the table cell's paragraph Range, sample every char's
vertical position (Information(6)), collect distinct Y levels (= line tops), report the
median consecutive delta = Word's per-font cell line height under docGrid linesAndChars
linePitch=350 + snapToGrid=0. Writes ASCII results to a file (cp932-safe)."""
import os, glob, json
import win32com.client as win32

REPRO = r'c:\tmp\cellrepro'
OUT = r'c:\tmp\cell_lineheight_word.json'
wdVertPos = 6

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0

manifest = {}
for line in open(os.path.join(REPRO, 'manifest.txt'), encoding='utf-8'):
    name, ea, sz = line.rstrip('\n').split('\t')
    manifest[name] = (ea, float(sz))

results = []
try:
    for path in sorted(glob.glob(os.path.join(REPRO, '*.docx'))):
        name = os.path.basename(path)
        ea, sz = manifest.get(name, ('?', 0))
        doc = word.Documents.Open(path, ReadOnly=True)
        try:
            tbl = doc.Tables(1)
            cell = tbl.Cell(1, 1)
            rng = cell.Range
            start, end = rng.Start, rng.End - 1  # exclude cell end marker
            # sample every char's Y
            ys = []
            i = start
            step = 1
            while i < end:
                cr = doc.Range(i, i)
                try:
                    ys.append(float(cr.Information(wdVertPos)))
                except Exception:
                    pass
                i += step
            # distinct line levels: round to 1pt, sorted unique
            levels = sorted(set(round(y) for y in ys))
            # merge levels within 3pt
            merged = []
            for y in levels:
                if merged and y - merged[-1] < 3:
                    continue
                merged.append(y)
            deltas = [round(merged[k] - merged[k - 1], 2) for k in range(1, len(merged))]
            # robust line pitch = median of deltas
            sd = sorted(deltas)
            med = sd[len(sd) // 2] if sd else None
            results.append({'name': name, 'font': ea, 'size_pt': sz,
                            'n_lines': len(merged), 'line_pitch_pt': med,
                            'deltas': deltas, 'levels': merged})
        finally:
            doc.Close(False)
finally:
    word.Quit()

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=1)

# ASCII summary
print("Word cell line pitch (docGrid linesAndChars linePitch=350=17.5pt, snapToGrid=0):")
print("font        size   n_lines  line_pitch(pt)   vs17.5   deltas(uniq)")
for r in results:
    ft = 'Mincho' if 'Mincho' in r['name'] else 'Gothic'
    uniq = sorted(set(r['deltas']))
    pitch = r['line_pitch_pt']
    vs = ('%+.2f' % (pitch - 17.5)) if pitch is not None else '  -  '
    print("  %-8s  %5.1f   %4d     %s        %s   %s" % (
        ft, r['size_pt'], r['n_lines'],
        ('%.2f' % pitch) if pitch is not None else ' - ', vs, uniq[:6]))
print("\nwrote", OUT)
