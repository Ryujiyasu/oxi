# -*- coding: utf-8 -*-
"""pcd bisection for BIG docs (R26 follow-up).

A full per-paragraph COM walk costs ~55 min on a 242-page doc (Repaginate
churn).  This samples ~N paragraphs with LONG unique text, reads their Word
page via Information(3) on a collapsed start (R30), matches each against the
Oxi --dump-layout by unique text prefix, and bisects the FIRST paragraph
where the page delta changes.  Minutes instead of an hour.

Usage:
    python _pcd_bisect.py <docx> <oxi_dump.json> [n_samples]

Output: the delta profile over the samples plus the bisected first-divergence
paragraph (index, texts, word page, oxi page).
"""
import os, re, sys, json, time
sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def norm(s):
    return re.sub(r"[\s\x07\r\x01\xa0]", "", s or "")


def build_oxi_index(dump_path):
    d = json.load(open(dump_path, encoding="utf-8"))
    occ = {}
    for pi, pg in enumerate(d["pages"], 1):
        groups = {}
        for e in pg["elements"]:
            t = e.get("text") or ""
            if not t.strip():
                continue
            key = (e.get("para_idx"), e.get("cell_row_idx"),
                   e.get("cell_col_idx"), e.get("cell_para_idx"))
            groups.setdefault(key, []).append((e["y"], e.get("x", 0), t))
        for key, frs in groups.items():
            frs.sort()
            y0 = frs[0][0]
            line1 = norm("".join(t for y, x, t in frs if abs(y - y0) < 1.5))
            if len(line1) >= 12:
                occ.setdefault(line1[:24], []).append(pi)
    return occ, len(d["pages"])


def main():
    docx, dump = sys.argv[1], sys.argv[2]
    n_samples = int(sys.argv[3]) if len(sys.argv) > 3 else 48
    occ, oxi_pages = build_oxi_index(dump)

    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        word.Options.UpdateLinksAtOpen = False
    except Exception:
        pass
    try:
        doc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True,
                                  AddToRecentFiles=False)
        try:
            n = doc.Paragraphs.Count
            print(f"paras {n}, oxi pages {oxi_pages}", flush=True)

            def probe(i):
                """(word_page, oxi_page, text) or None if unusable/ambiguous."""
                r = doc.Paragraphs(i).Range
                t = norm(r.Text)
                if len(t) < 12:
                    return None
                hits = occ.get(t[:24])
                if not hits or len(hits) != 1:
                    return None
                cr = doc.Range(r.Start, r.Start)
                return (cr.Information(3), hits[0], r.Text.strip()[:40])

            def delta_at(i, span=40):
                """Scan forward from i for the first usable paragraph."""
                for j in range(i, min(i + span, n + 1)):
                    p = probe(j)
                    if p:
                        return (j, p[1] - p[0], p)
                return None

            step = max(1, n // n_samples)
            samples = []
            for i in range(1, n + 1, step):
                s = delta_at(i)
                if s:
                    samples.append(s)
                    print(f"  i{s[0]:6d} w{s[2][0]:4d} o{s[1] + s[2][0] - s[1]:4d} "
                          f"d{s[1]:+d} {s[2][2]!r}", flush=True)
            # find first adjacent pair where delta changes
            for a, b in zip(samples, samples[1:]):
                if a[1] != b[1]:
                    lo, hi = a[0], b[0]
                    print(f"bisecting delta {a[1]:+d} -> {b[1]:+d} "
                          f"in i[{lo},{hi}]", flush=True)
                    dlo = a[1]
                    while hi - lo > 4:
                        mid = (lo + hi) // 2
                        s = delta_at(mid)
                        if s is None or s[0] >= hi:
                            hi = mid
                            continue
                        if s[1] == dlo:
                            lo = s[0]
                        else:
                            hi = s[0]
                        print(f"    i{s[0]} d{s[1]:+d}", flush=True)
                    print(f"FIRST DIVERGENCE between i{lo} and i{hi}")
                    for i in range(lo, min(hi + 1, n + 1)):
                        p = probe(i)
                        if p:
                            print(f"  i{i:6d} w{p[0]:4d} o{p[1]:4d} {p[2]!r}")
                    break
            else:
                print("no delta change across samples "
                      "(divergence after the last sample or non-monotone)")
        finally:
            doc.Close(False)
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
