# -*- coding: utf-8 -*-
"""Regenerate the JA blind Word pagination truth with TODAY's Word.

The stored truth (2026-08-21) wrapped a compat-11 line-final 。 that today's
Word hangs (educational__08709ff2 p5: 0.9466 -> 0.9924 once re-measured), so
every blind truth is suspect. Each old file is kept beside the new one as
`<did>.json.bak_20260821` (never overwritten if it already exists).

    python _ja_truth_refresh.py            # all 100 docs, both sets
    python _ja_truth_refresh.py 08709ff2   # only docs whose id contains the substring
"""
import json
import os
import shutil
import sys
import time
from pathlib import Path

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import measure_pagination_word as MW  # noqa: E402

REPO = Path(HERE).resolve().parents[1]
BENCH = REPO / "pipeline_data" / "ja_benchmark"
SETS = {"blind50": ("_final_jablind50.json", "p1_blind50"),
        "blindB50": ("_final_jablindB50.json", "p1_blindB50")}
only = sys.argv[1] if len(sys.argv) > 1 else None


def docs(setname):
    manifest, outdir = SETS[setname]
    data = json.loads((BENCH / manifest).read_text(encoding="utf-8"))
    for _t, lst in data.items():
        for c in lst:
            p = Path(c["path"])
            yield f"{p.parent.name}__{p.stem}", str(p.resolve()), BENCH / outdir


def main():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    n_done = n_changed = 0
    try:
        for setname in SETS:
            for did, path, outdir in docs(setname):
                if only and only not in did:
                    continue
                out = outdir / "word" / f"{did}.json"
                if not out.exists():
                    continue
                bak = out.with_suffix(".json.bak_20260821")
                if not bak.exists():
                    shutil.copyfile(out, bak)
                old = json.loads(out.read_text(encoding="utf-8"))
                t0 = time.time()
                try:
                    new = MW.measure_doc(app, path)
                except Exception as e:
                    print("  %s: FAILED %s" % (did, str(e)[:80]))
                    continue
                out.write_text(json.dumps(new, ensure_ascii=False, indent=1), encoding="utf-8")
                po = old.get("paragraphs", []) if isinstance(old, dict) else old
                pn = new.get("paragraphs", []) if isinstance(new, dict) else new
                diff = sum(1 for a, b in zip(po, pn) if a.get("page") != b.get("page")) + abs(len(po) - len(pn))
                n_done += 1
                if diff:
                    n_changed += 1
                print("  %-36s %5.1fs paras=%d page-changes=%d" % (did, time.time() - t0, len(pn), diff))
    finally:
        app.Quit()
    print("refreshed %d docs, %d with page changes" % (n_done, n_changed))


if __name__ == "__main__":
    main()
