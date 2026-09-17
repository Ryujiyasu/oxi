# -*- coding: utf-8 -*-
"""Corpus-scale census: which LINES does Word actually compress 約物 on, and by how much?

Every remaining golden markers-off failure (ikujidetail / nedocontract / ohnoshugyo, and the
tail of ohnoikuji) is the same shape: Oxi keeps one more character on a line than Word, so a
paragraph comes out one line short and a page swallows one extra paragraph. The per-line
credit is real -- disabling S475 (capacity) costs 32 documents' worth of paragraphs and
disabling S601 (hanging) costs 22 -- so the rule is right and the DISCRIMINATOR is missing.
Two single-document hypotheses have already failed:

  * S1462 "no compressible mark on the line -> no final-character tolerance" -- shipped, but
    only reaches a paragraph's LAST character.
  * S1463 "compress only when it saves a whole line" -- 0.9828 -> 0.8897 on ikujidetail.

So stop guessing per document and collect the ground truth per LINE instead.

Readout, per line of every multi-line body paragraph: the characters, each character's
advance (from Information(5) differences within the line), the plain-character modal
advance, every mark's advance against it, and the features a discriminator could use --
the line's index in its paragraph, whether it is the last line, the last character, the
first character of the NEXT line, and how far the line would overflow at natural widths.

`Information(5)` traps this instrument works around:
  * it is quantised to 0.75pt, so a single advance is unreliable -- read a run of them and
    take the mode;
  * for the FIRST character of a line it returns the LINE's left edge, not the glyph x, so
    the first advance of every line is dropped.

Output: one JSON object per line to stdout, ready to pipe into a JSONL file.
"""
import sys, os, json
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import win32com.client as win32

MARKS = '、。，．・：；（）「」［］〕〔【】'
DOCS = sys.argv[1:] or [
    'ikujidetail_002197815.docx',
    'ohnoshugyo_01.docx',
    'nedocontract_800052205.docx',
    'ohnoikuji_03.docx',
]
MAX_PARAS = int(os.environ.get('OIKOMI_MAX_PARAS', '120'))


def mode(vals):
    if not vals:
        return None
    best, n = None, -1
    for v in set(vals):
        c = vals.count(v)
        if c > n:
            best, n = v, c
    return best


app = win32.DispatchEx('Word.Application'); app.Visible = False
try:
    for name in DOCS:
        path = os.path.abspath(os.path.join('tools/golden-test/documents/docx', name))
        if not os.path.exists(path):
            print(json.dumps({'error': 'missing', 'doc': name}), flush=True)
            continue
        d = app.Documents.Open(path, ReadOnly=True)
        try:
            total = d.Paragraphs.Count
            done = 0
            for pi in range(1, total + 1):
                if done >= MAX_PARAS:
                    break
                rng = d.Paragraphs(pi).Range
                n = rng.Characters.Count
                if n < 45 or n > 400:
                    continue
                rows = []
                for k in range(1, n + 1):
                    c = rng.Characters(k)
                    rows.append((c.Text, c.Information(5), c.Information(6)))
                # split into lines on a y increase
                lines, cur, prev_y = [], [], None
                for ch, x, y in rows:
                    if prev_y is not None and y > prev_y + 1.0:
                        lines.append(cur); cur = []
                    if ch not in ('\r', '\x07'):
                        cur.append((ch, x))
                    prev_y = y
                if cur:
                    lines.append(cur)
                lines = [ln for ln in lines if ln]
                if len(lines) < 2:
                    continue
                done += 1
                for li, ln in enumerate(lines):
                    # advances inside the line; the FIRST character's x is the line edge
                    advs = [round(ln[i + 1][1] - ln[i][1], 2) for i in range(1, len(ln) - 1)]
                    plain = [a for (c, _), a in zip(ln[1:-1], advs) if c not in MARKS]
                    pm = mode(plain)
                    # each mark carries its LOCAL plain advance (the nearest plain
                    # neighbours) so a paragraph that mixes font sizes cannot inflate
                    # the apparent compression against a whole-line mode.
                    inner = list(zip([c for c, _ in ln[1:-1]], advs))
                    marks = []
                    for mi, (c, a) in enumerate(inner):
                        if c not in MARKS:
                            continue
                        near = [
                            advs[j]
                            for j in range(max(0, mi - 3), min(len(inner), mi + 4))
                            if inner[j][0] not in MARKS
                        ]
                        marks.append({'ch': c, 'adv': a, 'local': mode(near), 'at_end': mi == len(inner) - 1})
                    nxt = lines[li + 1][0][0] if li + 1 < len(lines) else None
                    print(json.dumps({
                        'doc': name,
                        'para': pi,
                        'line': li,
                        'n_lines': len(lines),
                        'is_last': li == len(lines) - 1,
                        'n_chars': len(ln),
                        'last_ch': ln[-1][0],
                        'next_ch': nxt,
                        'plain_mode': pm,
                        'x_first': ln[0][1],
                        'x_last': ln[-1][1],
                        'marks': marks,
                        'n_compressed': sum(
                            1 for m in marks if pm is not None and m['adv'] < pm - 0.8
                        ),
                    }, ensure_ascii=False), flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
