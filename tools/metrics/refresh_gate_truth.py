# -*- coding: utf-8 -*-
"""Re-export the Word pagination truth of every gate document on this machine.

The full gate (batch_gate / subset_gate: feature_census rows with a truth file in
pipeline_data/pagination_word or */p1_*/word) and the strict gate's manifest share
these files. Three of them moved between 2026-10-14 and 2026-10-27 on the same
PC, so the whole set is re-measured here: one Word process per document, nothing
else driving Word meanwhile (see memory word-truth-no-concurrent-com).

New truth goes to a staging directory first; `--install` copies it over the live
files after backing the old ones up. Resumable: documents already staged are
skipped.

    python tools/metrics/refresh_gate_truth.py measure [--limit N]
    python tools/metrics/refresh_gate_truth.py report
    python tools/metrics/refresh_gate_truth.py install
"""
import json, os, sys, shutil, subprocess, time
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
sys.path.insert(0, str(REPO / "pipeline_data" / "en_benchmark"))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
STAGE = Path(os.environ.get("OXI_TRUTH_STAGE", "E:/oxi-scratch/truth_refresh_20261027"))
BACKUP = STAGE / "_old"
MANIFEST = REPO / "pipeline_data/codex_phase1_dev_20260927/carrier-expanded1115-manifest.json"


def work_list():
    """[(truth_path, docx_path)] -- the union of the census gate and the strict manifest."""
    import feature_census as FC
    from subset_gate import truth_for
    pairs = {}
    rows = FC.load()
    for d, row in rows.items():
        t = truth_for(d)
        if t is not None and row.get("path"):
            pairs[str(Path(t).resolve())] = str(Path(row["path"]).resolve())
    if MANIFEST.exists():
        for case in json.load(open(MANIFEST, encoding="utf-8")):
            t, src = case.get("word"), case.get("input")
            if t and src and os.path.exists(t) and os.path.exists(src):
                pairs.setdefault(str(Path(t).resolve()), str(Path(src).resolve()))
    return sorted(pairs.items())


def staged(truth):
    rel = Path(truth).resolve().relative_to(REPO.resolve())
    return STAGE / rel


def measure(limit=None):
    import measure_pagination_word as MW
    import win32com.client
    STAGE.mkdir(parents=True, exist_ok=True)
    items = work_list()

    def old_pages(item):
        try:
            return json.load(open(item[0], encoding="utf-8")).get("n_pages") or 0
        except Exception:
            return 0
    # Light documents first: the 13 documents over 100 pages take hours between them.
    items.sort(key=old_pages)
    print(f"{len(items)} truth files", flush=True)
    log = open(STAGE / "_log.jsonl", "a", encoding="utf-8")
    done = 0
    for n, (truth, src) in enumerate(items, 1):
        out = staged(truth)
        if out.exists():
            continue
        if limit is not None and done >= limit:
            break
        out.parent.mkdir(parents=True, exist_ok=True)
        try:
            old = json.load(open(truth, encoding="utf-8"))
        except Exception:
            old = {}
        heavy = (old.get("n_pages") or 0) >= 100
        t0 = time.time()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            if n == 1 or not (STAGE / "_word_version.json").exists():
                json.dump({"version": word.Version, "build": word.Build}, open(STAGE / "_word_version.json", "w"))
            if heavy:
                import _fast_word_enum as FW
                res = FW.measure_doc_fast(word, src)
            else:
                res = MW.measure_doc(word, src)
            json.dump(res, open(out, "w", encoding="utf-8"))
            moved = sum(1 for a, b in zip(old.get("paragraphs", []), res.get("paragraphs", []))
                        if a.get("page") != b.get("page"))
            rec = {"truth": truth, "src": src, "old_pages": old.get("n_pages"), "new_pages": res.get("n_pages"),
                   "moved": moved, "n_paras_old": len(old.get("paragraphs", [])),
                   "n_paras_new": len(res.get("paragraphs", [])), "sec": round(time.time() - t0, 1)}
            print(f"  [{n:4}] {Path(truth).name}: {old.get('n_pages')}->{res.get('n_pages')}pg moved={moved} ({rec['sec']}s)", flush=True)
        except Exception as e:
            rec = {"truth": truth, "src": src, "error": str(e)[:200]}
            print(f"  [{n:4}] {Path(truth).name}: FAIL {str(e)[:120]}", flush=True)
        finally:
            try:
                word.Quit()
            except Exception:
                pass
        log.write(json.dumps(rec, ensure_ascii=False) + "\n"); log.flush()
        done += 1


def report():
    recs = {}
    for line in open(STAGE / "_log.jsonl", encoding="utf-8"):
        r = json.loads(line); recs[r["truth"]] = r
    ok = [r for r in recs.values() if "error" not in r]
    changed = [r for r in ok if r["moved"] or r["old_pages"] != r["new_pages"] or r["n_paras_old"] != r["n_paras_new"]]
    print(f"measured {len(ok)} / errors {len(recs) - len(ok)} / changed {len(changed)}")
    for r in sorted(changed, key=lambda r: -r["moved"])[:60]:
        print(f"  {Path(r['truth']).parent.parent.name}/{Path(r['truth']).name}: pages {r['old_pages']}->{r['new_pages']} moved {r['moved']} paras {r['n_paras_old']}->{r['n_paras_new']}")
    for r in recs.values():
        if "error" in r:
            print(f"  ERROR {Path(r['truth']).name}: {r['error'][:100]}")


def install():
    n = 0
    for truth, _src in work_list():
        new = staged(truth)
        if not new.exists():
            continue
        bk = BACKUP / Path(truth).resolve().relative_to(REPO.resolve())
        bk.parent.mkdir(parents=True, exist_ok=True)
        if not bk.exists():
            shutil.copy2(truth, bk)
        shutil.copy2(new, truth)
        n += 1
    print(f"installed {n} truth files (old copies under {BACKUP})")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "report"
    if cmd == "measure":
        lim = int(sys.argv[sys.argv.index("--limit") + 1]) if "--limit" in sys.argv else None
        measure(lim)
    elif cmd == "report":
        report()
    elif cmd == "install":
        install()
