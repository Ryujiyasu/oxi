"""S492 scoping — TRUE prevalence of wrapping non-justified paragraphs, via Word's
RESOLVED Format.Alignment (handles style inheritance my regex couldn't).

For a sample of real baseline docs, per paragraph: resolved alignment + whether it
wraps (>1 line, via Information(6) Y change across the range) + char count. Counts
how many WRAPPING paragraphs are non-justified (left/center/right) vs justified
(both/distribute). The non-justified wrapping ones are where Oxi's ungated
punct-compression break over-packs (S492 finding). Also dumps a few real
non-justified wrapping paragraphs' L1 char counts + whether Word compressed any
punct on L1 (the real-doc confirmation of 'jc!=both -> natural, zero compression').
"""
import glob, os, re
import win32com.client as w32

WD_VPOS = 6
ALIGN = {0: 'left', 1: 'center', 2: 'right', 3: 'both', 4: 'distribute', 5: 'thaiJ'}

# sample: bottom-N + spread across corpus
files = sorted(glob.glob('pipeline_data/golden_per_page/*_p1.docx'))
# prioritize known docs then sample the rest
prio = [f for f in files if re.search(r'(d77a|b837|ed025c|0e7a|683f|b35|15076|d4d126|c7b923|3a4f)', f)]
rest = [f for f in files if f not in prio]
sample = prio + rest[::6]  # ~all priority + every 6th of the rest
sample = sample[:40]

word = w32.DispatchEx('Word.Application'); word.Visible = False
from collections import Counter
wrap_align = Counter()      # alignment -> count of WRAPPING paragraphs
nonjust_examples = []
try:
    for f in sample:
        try:
            doc = word.Documents.Open(os.path.abspath(f), ReadOnly=True)
        except Exception:
            continue
        try:
            for p in doc.Paragraphs:
                rng = p.Range
                txt = rng.Text
                clean = txt.replace('\r', '').replace('\x07', '').replace('\n', '')
                if len(clean) < 20:
                    continue
                al = p.Alignment
                start, end = rng.Start, rng.End
                y0 = doc.Range(start, start).Information(WD_VPOS)
                yN = doc.Range(max(start, end - 1), max(start, end - 1)).Information(WD_VPOS)
                wraps = (yN - y0) > 2
                if not wraps:
                    continue
                aname = ALIGN.get(al, str(al))
                wrap_align[aname] += 1
                if aname in ('left', 'center', 'right') and len(nonjust_examples) < 12:
                    # measure L1 char count + count compressed punct on L1
                    n = 0
                    comp = 0
                    last_x = None
                    WD_HPOS = 5
                    advs = []
                    prev = None
                    for i in range(len(txt)):
                        ch = txt[i]
                        if ch in ('\r', '\n', '\x07'):
                            continue
                        x = doc.Range(start + i, start + i).Information(WD_HPOS)
                        if doc.Range(start + i, start + i).Information(WD_VPOS) > y0 + 2:
                            break
                        if prev is not None:
                            advs.append(round(x - prev[1], 2))
                        prev = (ch, x)
                        n += 1
                    # punct advances on L1 that are compressed (<11.0 for a 12pt fullwidth punct context)
                    nonjust_examples.append((os.path.basename(f)[:18], aname, n, advs[:14]))
        finally:
            doc.Close(False)
finally:
    word.Quit()

print("sampled docs:", len(sample))
print("WRAPPING paragraphs by resolved alignment:", dict(wrap_align))
total = sum(wrap_align.values())
nonjust = wrap_align['left'] + wrap_align['center'] + wrap_align['right']
print("wrapping total=%d ; non-justified wrapping=%d (%.0f%%)" % (total, nonjust, 100 * nonjust / max(1, total)))
print("\nReal non-justified WRAPPING paragraph L1 traces (advances; fullwidth natural=12.0):")
for stem, al, n, advs in nonjust_examples:
    nc = sum(1 for a in advs if 5 < a < 11.0)
    print("  %-18s %-6s L1=%2d compressed_in_first14=%d  advs=%s" % (stem, al, n, nc, advs))
