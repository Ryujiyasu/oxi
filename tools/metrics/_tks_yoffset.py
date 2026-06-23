# -*- coding: utf-8 -*-
"""Localize tokyoshugyo's cumulative vertical under-count: per-para absolute-Y
offset (Oxi - Word). Word (page,y,text) from pagination_word JSON (Information(6));
Oxi (page,y,text) from the --dump-layout. Match by normalized text prefix (robust,
monotonic). GROWTH points in the offset = where Oxi loses height vs Word.
"""
import json, glob, sys
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")

PH = 841.92  # A4 pt
WMAX = int(sys.argv[1]) if len(sys.argv) > 1 else 260


def norm(s):
    return (s or "").replace("　", "").replace(" ", "").replace("　", "").strip()


def word_paras():
    wf = glob.glob("pipeline_data/pagination_word/tokyoshugyo*.json")[0]
    W = json.load(open(wf, encoding="utf-8"))
    out = []
    for p in W["paragraphs"]:
        out.append({"i": p["i"], "page": p["page"], "y": p["y"],
                    "absY": (p["page"] - 1) * PH + p["y"], "t": norm(p["text"])})
    return out


def oxi_paras():
    d = json.load(open("C:/tmp/tks_dump.json", encoding="utf-8"))
    # collect first-line absY + text per (para_idx or cell) in document order
    rows = []
    seen = {}
    for pidx, pgd in enumerate(d["pages"]):
        # group by (para_idx, cell_para_idx) keeping min y
        grp = defaultdict(list)
        for e in pgd["elements"]:
            if e.get("type") != "text":
                continue
            key = (e.get("para_idx"), e.get("cell_para_idx"), e.get("cell_row_idx"), e.get("cell_col_idx"))
            grp[key].append(e)
        for key, es in grp.items():
            ys = defaultdict(list)
            for e in es:
                ys[round(e["y"], 1)].append(e)
            y0 = min(ys)
            txt = "".join(c["text"] for c in sorted(ys[y0], key=lambda c: c["x"]))
            rows.append({"page": pidx + 1, "absY": pidx * PH + y0, "t": norm(txt), "key": key})
    return rows


def main():
    wp = [w for w in word_paras() if w["i"] <= WMAX and w["t"]]
    op = [o for o in oxi_paras() if o["t"]]
    # monotonic match by prefix
    oi = 0
    prev_off = None
    print(f"{'wi':>5} {'Wpg':>3} {'Opg':>3} {'offset':>8} {'grow':>7}  text")
    for w in wp:
        pref = w["t"][:8]
        found = None
        for j in range(oi, min(oi + 60, len(op))):
            if op[j]["t"][:8] == pref or (len(pref) >= 5 and op[j]["t"].startswith(pref[:5])):
                found = j
                break
        if found is None:
            continue
        oi = found
        off = op[found]["absY"] - w["absY"]
        grow = "" if prev_off is None else f"{off - prev_off:+.1f}"
        mark = "  <<<" if prev_off is not None and abs(off - prev_off) > 3.0 else ""
        print(f"{w['i']:>5} {w['page']:>3} {op[found]['page']:>3} {off:>+8.1f} {grow:>7}{mark}  {w['t'][:30]}")
        prev_off = off


if __name__ == "__main__":
    main()
