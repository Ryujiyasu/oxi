"""Compat-15 justified shrink ALLOW as a function of the LINE'S LAST WORD.

Why: S1475 bounds the allowance by 0.35 x (last word + space), measured with
TNR-10 and letter words. Two real lines disagree with it in opposite
directions -- legal__0027c9c1 (Arial 11, «…October 11,»: Word shrinks 7.6pt
where the law gives 6.42) and reports__0013bcb8 (Book Antiqua 8pt w105,
«…semper vel,»: 4.88 against 4.73) -- and in both a comma or a period at the
end flips at the same width, so it is not a punctuation hang.

Method (render-measured, like _pb_jshrink_gen): HEAD = K-1 copies of
`mnopq`, then LAST, then filler words so the justified line 1 is not the
paragraph's last line.
  natural  = PDF bbox width of a LEFT-aligned line holding HEAD + LAST
  avail_min= smallest available width (1tw search) at which the JUSTIFIED doc
             keeps LAST on line 1
  allow    = natural - avail_min

    python tools/metrics/_pb_jlastword_gen.py [font_cfg ...]
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_jshrink_gen as J  # noqa: E402

K = 15                     # words on line 1 (14 spaces), as the legal line
PAGE = 16838               # A3-wide page so the justified line fits K words
FONTS = [
    ("arial11", "Arial", 22, None),
    ("tnr10", "Times New Roman", 20, None),
    ("ba8w105", "Book Antiqua", 16, 105),
]
LASTS = ["mnopq", "11,", "11.", "11", "October", "Oct,", "1,", "2021.", "mm,"]


def _rpr(font, sz, scale):
    sc = f'<w:w w:val="{scale}"/>' if scale else ""
    return (f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>{sc}'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')


def _para(cfg, last, justify, filler):
    _, font, sz, scale = cfg
    rpr = _rpr(font, sz, scale)
    words = ["mnopq"] * (K - 1) + [last] + (["mnopq"] * 30 if filler else [])
    text = " ".join(words)
    jc = "both" if justify else "left"
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:jc w:val="{jc}"/><w:rPr>{rpr}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


def measure(r, cfg, last):
    tag = f"lw_{cfg[0]}_{abs(hash(last)) % 10**6}"
    got = r.line1(J._document(_para(cfg, last, False, False), J.WIDE_PAGE_TW, 720), tag + "_n")
    if got is None or len(got[0].split()) != K:
        return None
    nat = got[2] - got[1]

    def keeps(right_tw):
        g = r.line1(J._document(_para(cfg, last, True, True), PAGE, right_tw), tag + f"_j{right_tw}")
        return g is not None and len(g[0].split()) >= K

    # right margin bracket: avail from nat+2 (keeps) down to nat-20 (wraps)
    lo = PAGE - J.LEFT_TW - int((nat + 2) * 20)   # small right margin -> keeps
    hi = PAGE - J.LEFT_TW - int((nat - 20) * 20)  # large right margin -> wraps
    if not keeps(lo) or keeps(hi):
        return nat, None
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if keeps(mid):
            lo = mid
        else:
            hi = mid
    return nat, (PAGE - J.LEFT_TW - lo) / 20.0


def main():
    names = sys.argv[1:]
    r = J.Renderer()
    try:
        for cfg in FONTS:
            if names and cfg[0] not in names:
                continue
            print(f"=== {cfg[0]}  K={K}")
            for last in LASTS:
                got = measure(r, cfg, last)
                if got is None or got[1] is None:
                    print(f"    {last!r:10} measure failed {got}")
                    continue
                nat, a = got
                print(f"    last={last!r:10} natural={nat:7.2f} avail_min={a:7.2f} allow={nat - a:6.2f}", flush=True)
    finally:
        r.close()


if __name__ == "__main__":
    main()
