# -*- coding: utf-8 -*-
"""Where does a document start losing (or gaining) height against Word?

`pagination_diff` answers at page granularity, which is too coarse once a document
is down to one flip: it says WHICH paragraph moved, not where the height went. This
walks the two side by side instead and prints, for every paragraph on the pages
asked for, the GAP to the previous paragraph in each engine. A gap that differs is
where a line was gained or lost; the constant offset between the two coordinate
systems cancels out of a gap, so nothing has to be calibrated.

Word's y is `Information(6)` = the cursor before the paragraph (memory:
feedback_text_y_vs_info6), Oxi's is the first line box's top, so compare GAPS only.

    python _pb_paragap.py <doc_id> <first_page> <last_page> [FLAG=1,FLAG2=1]
"""
import json
import os
import subprocess
import sys
import tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
DOCS = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
WORD = os.path.join(REPO, "pipeline_data", "pagination_word")
EXE = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")


def oxi_paras(docx, env_extra):
    env = dict(os.environ)
    env.update(env_extra)
    with tempfile.TemporaryDirectory() as td:
        out = os.path.join(td, "d.json")
        subprocess.run([EXE, docx, os.path.join(td, "p"), "96",
                        "--dump-layout=" + out], env=env, check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        d = json.load(open(out, encoding="utf-8"))
    out = []
    for pg in d["pages"]:
        best = {}
        for el in pg.get("elements", []):
            if el.get("type") != "text":
                continue
            key = (el.get("para_idx"), el.get("cell_para_idx"),
                   el.get("cell_row_idx"), el.get("cell_col_idx"))
            if key[0] is None:
                continue
            y = round(el["y"], 2)
            if key not in best or (y, el["x"]) < best[key][0]:
                best[key] = ((y, el["x"]), el.get("text", ""))
        for key, ((y, x), txt) in best.items():
            out.append((pg["page"], y, x, txt, key))
    out.sort(key=lambda r: (r[0], r[1], r[2]))
    return out


def norm(t):
    return "".join(c for c in t if not c.isspace())[:12]


def main():
    did, p0, p1 = sys.argv[1], int(sys.argv[2]), int(sys.argv[3])
    env = {}
    if len(sys.argv) > 4:
        for kv in sys.argv[4].split(","):
            k, _, v = kv.partition("=")
            env[k] = v or "1"
    docx = next(os.path.join(DOCS, f) for f in sorted(os.listdir(DOCS))
                if f.startswith(did) and f.endswith(".docx"))
    word = json.load(open(os.path.join(WORD, did + ".json"), encoding="utf-8"))
    wps = [p for p in word["paragraphs"] if p0 <= p.get("page", 0) <= p1
           and p.get("text", "").strip()]
    oxi = [r for r in oxi_paras(docx, env) if p0 <= r[0] <= p1 and r[3].strip()]
    print("%s  pages %d..%d   word paras %d / oxi %d   env=%s"
          % (did, p0, p1, len(wps), len(oxi), env or "default"))
    print("%-5s %-8s %-7s | %-5s %-8s %-7s | %-7s %s"
          % ("Wpg", "Wy", "Wgap", "Opg", "Oy", "Ogap", "d(gap)", "text"))
    j = 0
    wprev = oprev = None
    for w in wps:
        key = norm(w["text"])
        m = None
        for k in range(j, min(j + 12, len(oxi))):
            if norm(oxi[k][3]).startswith(key[:8]) or key.startswith(norm(oxi[k][3])[:8]):
                m = k
                break
        if m is None:
            print("%-5s %-8.2f %-7s | %-5s %-8s %-7s | %-7s %s"
                  % (w["page"], w["y"], "-", "-", "-", "-", "-", w["text"][:34]))
            continue
        o = oxi[m]
        wgap = (w["y"] - wprev["y"] + (w["page"] - wprev["page"]) * 1000.0
                if wprev else 0.0)
        ogap = (o[1] - oprev[1] + (o[0] - oprev[0]) * 1000.0) if oprev else 0.0
        d = ogap - wgap
        print("%-5d %-8.2f %-7.2f | %-5d %-8.2f %-7.2f | %-+7.2f %s%s"
              % (w["page"], w["y"], wgap, o[0], o[1], ogap, d, w["text"][:34],
                 "   <<<<" if abs(d) > 0.6 else ""))
        wprev, oprev = w, o
        j = m + 1


if __name__ == "__main__":
    main()
