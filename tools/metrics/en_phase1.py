# -*- coding: utf-8 -*-
"""Phase-1 (per-paragraph pagination vs Word COM) for the EN discovery
benchmark (docx-corpus 50). Reuses measure_pagination_word.measure_doc,
measure_pagination_oxi.measure_doc, and pagination_diff.diff_doc — the same
gate the JP corpus uses — on the frozen 50-doc selection (_final.json).

Phases (resumable; artifacts cached under pipeline_data/en_benchmark/p1/):
  python en_phase1.py word   # Word COM per-paragraph pages -> word/<id>.json
  python en_phase1.py oxi    # Oxi --dump-layout aggregate  -> oxi/<id>.json
  python en_phase1.py diff   # cross-join, per-doc pass/fail + report
  python en_phase1.py all
"""
import os, sys, json
from pathlib import Path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(r"c:\Users\ryuji\oxi-main")
BENCH = REPO / "pipeline_data" / "en_benchmark"
P1 = BENCH / "p1"
WDIR = P1 / "word"
ODIR = P1 / "oxi"


def selected():
    final = json.load(open(BENCH / "_final.json", encoding="utf-8"))
    docs = []
    for t, lst in final.items():
        for c in lst:
            did = f"{Path(c['path']).parent.name}__{Path(c['path']).stem}"
            docs.append((did, str(Path(c["path"]).resolve())))
    return docs


def phase_word():
    import measure_pagination_word as MW
    import win32com.client
    WDIR.mkdir(parents=True, exist_ok=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    for did, path in selected():
        out = WDIR / f"{did}.json"
        if out.exists():
            continue
        try:
            res = MW.measure_doc(word, path)
            json.dump(res, open(out, "w", encoding="utf-8"))
            print(f"  word {did}: {res.get('n_pages')}pg {len(res.get('paragraphs', res.get('rows', [])))}para")
        except Exception as e:
            print(f"  word {did}: FAIL {str(e)[:60]}")
    word.Quit()


def phase_oxi():
    import measure_pagination_oxi as MO
    ODIR.mkdir(parents=True, exist_ok=True)
    for did, path in selected():
        out = ODIR / f"{did}.json"
        if out.exists():
            continue
        try:
            res = MO.measure_doc(path)
            json.dump(res, open(out, "w", encoding="utf-8"))
            print(f"  oxi {did}: {res.get('n_pages')}pg")
        except Exception as e:
            print(f"  oxi {did}: FAIL {str(e)[:60]}")


def phase_diff():
    import pagination_diff as PD
    rows = []
    for did, _ in selected():
        wf = WDIR / f"{did}.json"; of = ODIR / f"{did}.json"
        if not wf.exists() or not of.exists():
            rows.append({"doc": did, "pass": None, "score": None, "pcd": None})
            continue
        word = json.load(open(wf, encoding="utf-8"))
        oxi = json.load(open(of, encoding="utf-8"))
        r = PD.diff_doc(did, word, oxi)
        rows.append({"doc": did, "pass": r["pass"], "score": r["score"],
                     "pcd": r["page_count_delta"], "hist": r["delta_histogram"]})
    json.dump(rows, open(P1 / "_result.json", "w", encoding="utf-8"), indent=1)
    scored = [r for r in rows if r["pass"] is not None]
    npass = sum(1 for r in scored if r["pass"])
    print(f"\n=== EN discovery Phase-1 ({len(scored)} docs measured) ===")
    print(f"pagination PASS: {npass}/{len(scored)}  ({100*npass/max(len(scored),1):.0f}%)")
    print(f"mean score: {sum(r['score'] for r in scored)/max(len(scored),1):.4f}")
    print("\nFAILs (worst pcd first):")
    for r in sorted([r for r in scored if not r["pass"]],
                    key=lambda x: -abs(x["pcd"] or 0)):
        print(f"  pcd={r['pcd']:+d} score={r['score']:.3f}  {r['doc']}  {r.get('hist')}")


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "all"
    if mode in ("word", "all"): phase_word()
    if mode in ("oxi", "all"): phase_oxi()
    if mode in ("diff", "all"): phase_diff()
