"""Scan baseline docs for <wp:anchor> + <wp:positionV>/<wp:positionH>
combinations. Identify which (relativeFrom × align/posOffset) pairs are
used in real documents, to prioritize which sub-cases to investigate.
"""
import re
import zipfile
from pathlib import Path
from collections import Counter

DOCX_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx")


# Both wp:anchor and v:shape (VML) carry positioning. Focus on wp: form first.
ANCHOR_RE = re.compile(r'<wp:anchor\b[^>]*>.*?</wp:anchor>', re.S)
POS_V_RE = re.compile(
    r'<wp:positionV\s+relativeFrom="([^"]*)">'
    r'(?:\s*<wp:posOffset>(-?\d+)</wp:posOffset>'
    r'|\s*<wp:align>([^<]*)</wp:align>'
    r')',
    re.S,
)
POS_H_RE = re.compile(
    r'<wp:positionH\s+relativeFrom="([^"]*)">'
    r'(?:\s*<wp:posOffset>(-?\d+)</wp:posOffset>'
    r'|\s*<wp:align>([^<]*)</wp:align>'
    r')',
    re.S,
)


def scan(p):
    try:
        with zipfile.ZipFile(p) as z:
            try:
                xml = z.read("word/document.xml").decode("utf-8")
            except KeyError:
                return None
    except Exception:
        return None
    anchors = ANCHOR_RE.findall(xml)
    if not anchors:
        return None
    summaries = []
    for a in anchors:
        v = POS_V_RE.search(a)
        h = POS_H_RE.search(a)
        v_rel = v.group(1) if v else None
        v_kind = "offset" if v and v.group(2) is not None else ("align" if v else "none")
        v_val  = v.group(2) if v and v.group(2) is not None else (v.group(3) if v else None)
        h_rel = h.group(1) if h else None
        h_kind = "offset" if h and h.group(2) is not None else ("align" if h else "none")
        h_val  = h.group(2) if h and h.group(2) is not None else (h.group(3) if h else None)
        # behindDoc / layoutInCell / allowOverlap flags
        flags = []
        for f in ["behindDoc", "layoutInCell", "allowOverlap", "locked"]:
            m = re.search(rf'\b{f}="(\d)"', a)
            if m and m.group(1) == "1":
                flags.append(f)
        wrap = None
        for w in ["wrapNone","wrapSquare","wrapTight","wrapTopAndBottom","wrapThrough"]:
            if f"<wp:{w}" in a:
                wrap = w
                break
        summaries.append({
            "v_rel": v_rel, "v_kind": v_kind, "v_val": v_val,
            "h_rel": h_rel, "h_kind": h_kind, "h_val": h_val,
            "flags": flags,
            "wrap": wrap,
        })
    return summaries


def main():
    docs = sorted(DOCX_DIR.glob("*.docx"))
    n_with = 0
    total_anchors = 0
    v_rel_counter = Counter()
    h_rel_counter = Counter()
    wrap_counter = Counter()
    flag_counter = Counter()
    pair_counter = Counter()  # (v_rel, h_rel)
    docs_with_anchor = []
    for p in docs:
        s = scan(p)
        if not s:
            continue
        n_with += 1
        total_anchors += len(s)
        docs_with_anchor.append((p.name, len(s)))
        for a in s:
            v_rel_counter[a["v_rel"]] += 1
            h_rel_counter[a["h_rel"]] += 1
            wrap_counter[a["wrap"]] += 1
            for f in a["flags"]:
                flag_counter[f] += 1
            pair_counter[(a["v_rel"], a["h_rel"])] += 1

    print(f"Scanned {len(docs)} docs; {n_with} with <wp:anchor> (total {total_anchors} anchors)")
    print()
    print("=== positionV relativeFrom counts ===")
    for k, v in sorted(v_rel_counter.items(), key=lambda x: -x[1]):
        print(f"  {k!r:30s} {v}")
    print()
    print("=== positionH relativeFrom counts ===")
    for k, v in sorted(h_rel_counter.items(), key=lambda x: -x[1]):
        print(f"  {k!r:30s} {v}")
    print()
    print("=== wrap mode counts ===")
    for k, v in sorted(wrap_counter.items(), key=lambda x: -x[1]):
        print(f"  {k!r:30s} {v}")
    print()
    print("=== flag counts (set=1 only) ===")
    for k, v in sorted(flag_counter.items(), key=lambda x: -x[1]):
        print(f"  {k!r:30s} {v}")
    print()
    print("=== (v_rel, h_rel) pairs ===")
    for (vr, hr), n in sorted(pair_counter.items(), key=lambda x: -x[1])[:20]:
        print(f"  ({vr!r:14s}, {hr!r:14s}) {n}")
    print()
    print("=== top docs with <wp:anchor> ===")
    for name, n in sorted(docs_with_anchor, key=lambda x: -x[1])[:10]:
        print(f"  {n:3d}  {name}")


if __name__ == "__main__":
    main()
