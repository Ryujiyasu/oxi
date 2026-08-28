# -*- coding: utf-8 -*-
"""Order-aware Word-vs-Oxi page alignment for a pcd!=0 doc.

Unique-text matching (the _pcd_bisect approach) breaks down on documents that
REPEAT text -- a legal rules book repeats form headers dozens of times, so most
paragraphs are discarded as ambiguous and the few that survive mis-pair, showing
fake deltas of +24 or -473.

This aligns the two paragraph SEQUENCES with difflib instead, so duplicates are
resolved by position, then reports where the page delta changes.

Order-awareness is not enough on its own. Where a long region fails to align at
all, a SHORT matching block on the far side of the gap can still pair the wrong
repeat and monotonicity never notices. legal__0010437a7f75f636 reported
`wi=8593 w=625 o=619 delta=-6  Citation` from a 3-long block, across a
395-paragraph blind stretch; both engines in fact spend 8 pages there, and the
document repeats that table header 8 times. So:

  * a delta transition is printed SUSPECT when it jumps by more than one page, or
    when the two pairs around it are separated by a long unaligned stretch --
    real drift moves a page at a time and is seen moving;
  * the blind regions are printed, so an unobserved stretch reads as unobserved
    instead of as a finding.

Dropping short matching blocks does NOT work as a guard and was tried: at
--minblk=5 this document keeps 69 of 1297 pairs and then reports 0 misplaced
paragraphs, which is worse than the artifact. Oxi emits more paragraph-groups
than Word has paragraphs, so trustworthy matches are routinely short.

Usage: _pcd_align.py <word_pagination.json> <oxi_dump.json> [--minblk=N]
"""
import json, re, sys, difflib
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

def norm(s):
    s = re.sub(r"[.]{3,}", "", s or "")
    return re.sub(r"[\s\x07\r\x01\xa0]", "", s)

args = [a for a in sys.argv[1:] if not a.startswith("--")]
MINBLK = next((int(a.split("=", 1)[1]) for a in sys.argv[1:]
               if a.startswith("--minblk=")), 1)
# A transition whose two pairs sit this far apart in Word paragraphs has an
# unobserved stretch between them and cannot be attributed to a site.
GAP_SUSPECT = 40
wpath, opath = args[0], args[1]
w = json.load(open(wpath, encoding="utf-8"))
wp = w["paragraphs"]
o = json.load(open(opath, encoding="utf-8"))["pages"]

# Word side: ordered (page, text)
wseq = [(p["page"], norm(p.get("text"))) for p in wp]

# Oxi side: ordered (page, first-line text) in reading order
oseq = []
for pi, pg in enumerate(o, 1):
    groups = {}
    for e in pg["elements"]:
        t = e.get("text") or ""
        if not t.strip():
            continue
        key = (e.get("para_idx"), e.get("cell_row_idx"), e.get("cell_col_idx"),
               e.get("cell_para_idx"))
        groups.setdefault(key, []).append((e["y"], e.get("x", 0), t))
    items = []
    for key, frs in groups.items():
        frs.sort()
        items.append((frs[0][0], frs[0][1], norm("".join(t for _, _, t in frs))))
    items.sort()
    for _, _, t in items:
        oseq.append((pi, t))

print(f"word paragraphs={len(wseq)}  oxi paragraph-groups={len(oseq)}")
print(f"word pages={w['n_pages']}  oxi pages={len(o)}  pcd={len(o)-w['n_pages']:+d}")

a = [t for _, t in wseq]
b = [t for _, t in oseq]
sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
pairs = []
dropped = 0
for blk in sm.get_matching_blocks():
    if blk.size < MINBLK:
        dropped += blk.size
        continue
    for k in range(blk.size):
        wi, oi = blk.a + k, blk.b + k
        if len(a[wi]) >= 8:
            pairs.append((wi, wseq[wi][0], oseq[oi][0], a[wi]))
print(f"aligned pairs (len>=8, block>={MINBLK}): {len(pairs)}  "
      f"({100*len(pairs)/max(1,len(wseq)):.0f}% of Word paragraphs); "
      f"{dropped} matches in short blocks dropped")

prev = None
prev_wi = None
print("\ndelta transitions (word_para_idx, word_page, oxi_page, delta):")
for wi, wpg, opg, t in pairs:
    d = opg - wpg
    if d != prev:
        # Real drift moves a page at a time. A bigger step means the pairs on
        # either side straddle a stretch nothing aligned in -- read the blind
        # regions below before treating it as a finding.
        why = []
        if prev is not None and abs(d - prev) > 1:
            why.append("jump>1")
        if prev_wi is not None and wi - prev_wi > GAP_SUSPECT:
            why.append(f"{wi - prev_wi} paras unaligned before it")
        sus = ("  SUSPECT (" + ", ".join(why) + ")") if why else ""
        print(f"   wi={wi:5d} w={wpg:4d} o={opg:4d} delta={d:+3d}   {t[:46]}{sus}")
        prev = d
    prev_wi = wi

# Where the alignment is blind: consecutive Word paragraphs with no trusted pair.
if pairs:
    blind = []
    idx = [wi for wi, _, _, _ in pairs]
    for lo, hi in zip(idx, idx[1:]):
        if hi - lo > 1:
            blind.append((lo + 1, hi - 1, wseq[lo][0], wseq[hi][0]))
    blind.sort(key=lambda r: r[3] - r[2], reverse=True)
    print("\nblind regions (no trusted pair), widest by Word pages spanned:")
    for lo, hi, wp0, wp1 in blind[:6]:
        print(f"   wi {lo}..{hi} ({hi - lo + 1} paras)  Word p{wp0}..p{wp1}"
              f"  = {wp1 - wp0} pages unobserved")
nz = sum(1 for _, wpg, opg, _ in pairs if opg != wpg)
print(f"\nparagraphs on the wrong page: {nz}/{len(pairs)} ({100*nz/max(1,len(pairs)):.1f}%)")
