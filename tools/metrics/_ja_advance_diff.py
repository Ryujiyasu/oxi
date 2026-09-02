# -*- coding: utf-8 -*-
"""Per-paragraph vertical ADVANCE, Word vs Oxi, for a JA benchmark doc.

A doc whose pages hold fewer paragraphs than Word's is thicker somewhere; this
says WHERE. It compares the advance between consecutive paragraphs, never the
absolute y: Oxi's element y is `cursor_before + text_y_off` while Word's
Information(6) is `cursor_before` alone, so absolute y carries a constant
offset that the difference cancels ([[feedback_text_y_vs_info6]]).

Paragraph sequences are aligned with difflib, so repeated text pairs by
position rather than by a lucky unique prefix.

    python _ja_advance_diff.py <set> <doc-substring> [--page N] [--eps 0.4]
"""
import difflib
import json
import os
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
REPO = Path(__file__).resolve().parents[2]
BENCH = REPO / "pipeline_data" / "ja_benchmark"
SETS = {"blind50": ("_final_jablind50.json", "p1_blind50"),
        "blindB50": ("_final_jablindB50.json", "p1_blindB50")}


def dec(t):
    try:
        return t.encode("latin-1").decode("cp932")
    except Exception:
        return t


def norm(s):
    return re.sub(r"\s+", "", dec(s or ""))[:24]


def main():
    setname, needle = sys.argv[1], sys.argv[2]
    want_page = None
    eps = 0.4
    for i, a in enumerate(sys.argv):
        if a == "--page":
            want_page = int(sys.argv[i + 1])
        if a == "--eps":
            eps = float(sys.argv[i + 1])

    import measure_pagination_oxi as MO

    manifest, outdir = SETS[setname]
    data = json.loads((BENCH / manifest).read_text(encoding="utf-8"))
    target = None
    for _t, lst in data.items():
        for c in lst:
            p = Path(c["path"])
            did = f"{p.parent.name}__{p.stem}"
            if needle in did:
                target = (did, str(p.resolve()), BENCH / outdir)
    if not target:
        print("no doc matched")
        return
    did, path, od = target
    word = json.loads((od / "word" / f"{did}.json").read_text(encoding="utf-8"))
    oxi = MO.measure_doc(path)

    wp = [p for p in word["paragraphs"]]
    op = []
    for pg_str, recs in sorted(oxi["pages"].items(), key=lambda kv: int(kv[0])):
        for r in recs:
            op.append({"page": int(pg_str), "y": r.get("y"), "text": r.get("text")})

    ws = [norm(p["text"]) for p in wp]
    os_ = [norm(p["text"]) for p in op]
    sm = difflib.SequenceMatcher(a=ws, b=os_, autojunk=False)
    pairs = []
    for a, b, n in sm.get_matching_blocks():
        for k in range(n):
            pairs.append((wp[a + k], op[b + k]))
    print(f"{did}: word {word['n_pages']}pg  oxi {oxi['n_pages']}pg  "
          f"aligned {len(pairs)}/{len(wp)} paragraphs")

    print(f"\n{'wpg':>4} {'opg':>4} {'w_adv':>8} {'o_adv':>8} {'diff':>8}   text")
    cum = 0.0
    for i in range(len(pairs) - 1):
        w0, o0 = pairs[i]
        w1, o1 = pairs[i + 1]
        # Only compare advances INSIDE a page on both sides; a page break makes
        # the difference meaningless (the cursor resets to the top margin).
        if w0["page"] != w1["page"] or o0["page"] != o1["page"]:
            cum = 0.0
            continue
        if want_page is not None and w0["page"] != want_page:
            continue
        if w0["y"] is None or o0["y"] is None or w1["y"] is None or o1["y"] is None:
            continue
        wa = w1["y"] - w0["y"]
        oa = o1["y"] - o0["y"]
        d = oa - wa
        cum += d
        if abs(d) >= eps:
            print(f"{w0['page']:>4} {o0['page']:>4} {wa:>8.2f} {oa:>8.2f} {d:>+8.2f}"
                  f"   {dec(w0['text'])[:34]!r}")
    if want_page is not None:
        print(f"\ncumulative excess on word page {want_page}: {cum:+.2f}pt")


if __name__ == "__main__":
    main()
