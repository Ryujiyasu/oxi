# -*- coding: utf-8 -*-
"""S492-1ec1 (S469 dedicated): robust COM measurement of Word's INTERNAL textbox line pitch.
PNG peak-detection failed in 1ec1's overlapping region; use Word Shape.TextFrame.TextRange +
Information(6) per character to get the rendered line Ys INSIDE each textbox. Find the blk3 box
(text contains the heading needle), sample char Ys, compute internal line Ys + pitches, compare
to Oxi (10.5pt content pitch ~17.0, glyph Ys 178.9/199.4/216.4/238.05/260.8/277.9). cp932-safe:
this file is UTF-8 (Write-authored); needle is a literal here (safe in a UTF-8 file, NOT a bash
heredoc); results to a JSON + ASCII summary."""
import json, glob
import win32com.client as win32

DOCX = glob.glob(r'c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\1ec1091177b1*.docx')[0]
OUT = r'c:\tmp\1ec1_tbx_pitch.json'
NEEDLE = '納税者に納税額'  # blk3 heading
wdVertPos = 6

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
result = {'shapes': [], 'target': None}
try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    for si in range(1, doc.Shapes.Count + 1):
        sh = doc.Shapes(si)
        txt = ''
        try:
            if sh.TextFrame.HasText:
                txt = sh.TextFrame.TextRange.Text or ''
        except Exception:
            pass
        # sanity: confirm needle membership works (cp932 trap guard)
        has = NEEDLE in txt
        result['shapes'].append({'i': si, 'has_needle': has, 'textlen': len(txt), 'text_head': txt[:20]})
        if has and result['target'] is None:
            rng = sh.TextFrame.TextRange
            # Textbox content is a SEPARATE story; address via the shape's own
            # TextRange.Characters (NOT doc.Range which maps to the body story).
            ncha = rng.Characters.Count
            ys = []
            step = max(1, ncha // 120)
            k = 1
            while k <= ncha:
                try:
                    ys.append(float(rng.Characters(k).Information(wdVertPos)))
                except Exception:
                    pass
                k += step
            # distinct line Ys
            levels = sorted(set(round(y, 1) for y in ys))
            merged = []
            for y in levels:
                if merged and y - merged[-1] < 3:
                    continue
                merged.append(y)
            pitches = [round(merged[k] - merged[k - 1], 2) for k in range(1, len(merged))]
            result['target'] = {'shape_i': si, 'line_ys': merged, 'pitches': pitches,
                                'top': sh.Top, 'height': sh.Height, 'n_samples': len(ys)}
    doc.Close(False)
finally:
    word.Quit()

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=1)

print("1ec1 shapes with text: %d" % sum(1 for s in result['shapes'] if s['textlen'] > 0))
print("needle-matched shapes: %d (cp932 guard: if 0, membership failed)" % sum(1 for s in result['shapes'] if s['has_needle']))
t = result['target']
if t:
    print("\nblk3 textbox (Shape %d) Word INTERNAL line Ys:" % t['shape_i'])
    print("  Shape.Top=%.1f Height=%.1f n_samples=%d" % (t['top'], t['height'], t['n_samples']))
    print("  line Ys: %s" % [round(y, 1) for y in t['line_ys']])
    print("  PITCHES: %s" % t['pitches'])
    print("\n  Oxi blk3 glyph Ys: 178.9(14) 199.4 216.4(10.5) 238.05(14) 260.8 277.9 ; 10.5 pitch ~17.0")
    print("  => compare Word internal pitches to Oxi ~17 to find the line-height/centering mismatch")
print("wrote", OUT)
