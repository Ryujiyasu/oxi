# -*- coding: utf-8 -*-
"""Full layout dumps of the whole census corpus, for refactor identity checks.

  python dump_corpus.py run <tag> <renderer.exe> [--jobs 3]
      -> pipeline_data/layout_dumps/<tag>/<doc_key>.json  (renderer's --dump-layout,
         written by the renderer itself, so bytes are exactly what it emits)
  python dump_corpus.py compare <tagA> <tagB>
      -> lists documents whose dumps differ (or exist on one side only)

A code change meant to be behaviour-preserving (splitting a function, moving
code between files) must leave every dump byte-identical; the pagination gate
alone would miss a sub-point drift that happens not to move a page break.
"""
import os, sys, json, subprocess, tempfile, filecmp
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = REPO / "pipeline_data" / "layout_dumps"
CHUNK = 24


def key_of(doc):
    return doc.replace("/", "__").replace("\\", "__")


def run_chunk(exe, items, outdir):
    """items: [(doc_key, docx_path)] -> set of keys the renderer reported OK."""
    ok = set()
    with tempfile.TemporaryDirectory(prefix="dc_") as tmp:
        lst = os.path.join(tmp, "list.txt")
        with open(lst, "w", encoding="utf-8") as f:
            for k, p in items:
                f.write(f"{p}\t{outdir / (k + '.json')}\n")
        try:
            r = subprocess.run([exe, "--batch=" + lst], capture_output=True, text=True,
                               encoding="utf-8", errors="replace", timeout=240 * len(items))
        except subprocess.TimeoutExpired:
            return ok
        for line in r.stdout.splitlines():
            if line.startswith("BATCH-OK "):
                path = line[len("BATCH-OK "):].strip()
                for k, p in items:
                    if p == path:
                        ok.add(k)
    return ok


def run(tag, exe, jobs):
    import feature_census as FC
    rows = FC.load()
    outdir = OUT / tag
    outdir.mkdir(parents=True, exist_ok=True)
    items = [(key_of(d), rows[d]["path"]) for d in sorted(rows)]
    items = [it for it in items if not (outdir / (it[0] + ".json")).exists()]
    print(f"{tag}: {len(rows)} docs, {len(items)} to render, jobs={jobs}", flush=True)
    chunks = [items[i:i + CHUNK] for i in range(0, len(items), CHUNK)]
    n_ok = 0
    with ThreadPoolExecutor(jobs) as pool:
        for i, ok in enumerate(pool.map(lambda c: run_chunk(exe, c, outdir), chunks)):
            n_ok += len(ok)
            if i % 10 == 9:
                print(f"  {min((i + 1) * CHUNK, len(items))}/{len(items)}", flush=True)
    have = len(list(outdir.glob("*.json")))
    print(f"done: {n_ok} rendered this run, {have}/{len(rows)} dumps present")
    missing = [k for k, _ in [(key_of(d), None) for d in rows] if not (outdir / (k + ".json")).exists()]
    for k in missing[:20]:
        print("  missing", k)


def compare(a, b):
    da, db = OUT / a, OUT / b
    ka = {p.name for p in da.glob("*.json")}
    kb = {p.name for p in db.glob("*.json")}
    only_a, only_b = sorted(ka - kb), sorted(kb - ka)
    diff = sorted(n for n in ka & kb if not filecmp.cmp(da / n, db / n, shallow=False))
    print(f"{a} vs {b}: {len(ka & kb)} common, {len(diff)} differ, {len(only_a)} only in {a}, {len(only_b)} only in {b}")
    for n in diff[:40]:
        print("  DIFFERS", n)
    for n in only_a[:10]:
        print(f"  only {a}:", n)
    for n in only_b[:10]:
        print(f"  only {b}:", n)
    return len(diff) + len(only_a) + len(only_b)


if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "run":
        a = sys.argv[2:]
        jobs = int(a[a.index("--jobs") + 1]) if "--jobs" in a else 3
        run(a[0], os.path.abspath(a[1]), jobs)
    elif cmd == "compare":
        sys.exit(1 if compare(sys.argv[2], sys.argv[3]) else 0)
