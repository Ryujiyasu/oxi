# -*- coding: utf-8 -*-
"""Batched iteration gate: several fixes, one run, automatic attribution.

  python batch_gate.py "<census expr>" --exe NEW.exe --base BASE.exe
         [--flags OXI_S1589_DISABLE,OXI_S1590_DISABLE] [--jobs 3] [--identity 12]

What it does, compared with subset_gate.py:
  * runs documents in PARALLEL: --jobs renderer processes, each laying out a
    chunk of documents in ONE process (`--batch`; a cold start is ~1.25s of a
    ~1.6s median document, so per-document processes spent most of the gate
    starting up); each renderer is capped at OXI_MEM_CAP_MB, so the default
    3 fits a 14GB machine;
  * CACHES every binary's result per document under
    pipeline_data/gate_cache/<sha12 of the binary>/ -- the base AND the new
    binary (when run without --new-env), so a binary is measured once, ever,
    and the commit you gate today is tomorrow's cached base;
  * for every PASS->FAIL document, re-runs that document once per --flags entry
    (the fix's opt-out env) and reports which flag restores the PASS, so a batch
    of fixes needs one gate run and at most (#regressions x #flags) re-renders.

Full gates remain the commit rule; this is the fast loop.
"""
import os, sys, json, random, subprocess, hashlib
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import feature_census as FC  # noqa: E402
from subset_gate import truth_for  # noqa: E402

CACHE = REPO / "pipeline_data" / "gate_cache"

CHUNK = 24


def _run_chunk(exe, items, extra_env):
    """items: [(key, path, truth)] -> {key: result} with one batch renderer."""
    import measure_pagination_oxi as MO, pagination_diff as PD
    env = dict(os.environ)
    env.update(extra_env or {})
    env["OXI_GDI_EXE"] = exe
    outs = MO.measure_docs_batch([it[1] for it in items], exe=exe, env=env)
    res = {}
    for (key, _path, truth), o in zip(items, outs):
        if isinstance(o, Exception):
            res[key] = {"pass": None, "err": str(o)[-200:]}
            continue
        try:
            word = json.load(open(truth, encoding="utf-8"))
            d = PD.diff_doc("x", word, o)
            res[key] = {"pass": d["pass"], "score": d["score"], "pcd": d.get("page_count_delta")}
        except Exception as e:
            res[key] = {"pass": None, "err": str(e)[-200:]}
    return res


def run_many(tasks, jobs):
    """tasks: [(key, exe, path, truth, env)] -> {key: result}. Documents that
    share a binary and env go through batch renderers CHUNK at a time."""
    groups = {}
    for key, exe, path, truth, env in tasks:
        groups.setdefault((exe, tuple(sorted((env or {}).items()))), []).append((key, path, str(truth)))
    chunks = [(exe, items[i:i + CHUNK], dict(envt))
              for (exe, envt), items in groups.items() for i in range(0, len(items), CHUNK)]
    out = {}
    total = sum(len(c[1]) for c in chunks)
    with ThreadPoolExecutor(max_workers=jobs) as pool:
        for r in pool.map(lambda c: _run_chunk(*c), chunks):
            out.update(r)
            print(f"  progress {len(out)}/{total}", flush=True)
    return out


def sha12(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()[:12]


def dump(exe, path):
    import tempfile, shutil
    tmp = tempfile.mkdtemp(prefix="bg_")
    try:
        out = os.path.join(tmp, "l.json")
        subprocess.run([exe, path, os.path.join(tmp, "p"), "110", "--dump-layout=" + out], capture_output=True)
        return open(out, encoding="utf-8").read() if os.path.exists(out) else None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    expr = sys.argv[1]
    a = sys.argv[2:]
    arg = lambda k, d=None: a[a.index(k) + 1] if k in a else d
    exe, base = arg("--exe"), arg("--base")
    flags = [f for f in (arg("--flags", "") or "").split(",") if f]
    jobs = int(arg("--jobs", "3"))
    n_id = int(arg("--identity", "0"))
    new_env = dict(kv.split("=", 1) for kv in (arg("--new-env", "") or "").split(",") if kv)
    rows = FC.load()
    hits = FC.query(expr)
    work = [(d, rows[d]["path"], truth_for(d)) for d in hits]
    work = [w for w in work if w[2] is not None]
    print(f"predicate: {expr}\nmatched {len(hits)} docs, {len(work)} with truth, jobs={jobs}", flush=True)

    def cached(binary, env):
        d = CACHE / sha12(binary)
        d.mkdir(parents=True, exist_ok=True)
        got, todo = {}, []
        for w in work:
            f = d / (w[0].replace("/", "__") + ".json")
            if not env and f.exists():
                got[w[0]] = json.loads(f.read_text(encoding="utf-8"))
            else:
                todo.append(w)
        return d, got, todo

    bdir, base_res, todo_base = cached(base, None)
    ndir, new_res, todo_new = cached(exe, new_env)
    print(f"cached: base {len(base_res)}/{len(work)}  new {len(new_res)}/{len(work)}", flush=True)
    tasks = ([(("b", w[0]), base, w[1], w[2], None) for w in todo_base]
             + [(("n", w[0]), exe, w[1], w[2], new_env or None) for w in todo_new])
    for (side, key), r in run_many(tasks, jobs).items():
        (base_res if side == "b" else new_res)[key] = r
        if r.get("pass") is not None and (side == "b" or not new_env):
            ((bdir if side == "b" else ndir) / (key.replace("/", "__") + ".json")).write_text(json.dumps(r), encoding="utf-8")
    base_res = [base_res[w[0]] for w in work]
    new_res = [new_res[w[0]] for w in work]

    flips = []
    n_new = n_base = 0
    for w, n, b in zip(work, new_res, base_res):
        n_new += n.get("pass") is True
        n_base += b.get("pass") is True
        tag = ""
        if b.get("pass") is True and n.get("pass") is not True:
            tag = "  <<< PASS->FAIL"; flips.append(w)
        elif b.get("pass") is not True and n.get("pass") is True:
            tag = "  >>> FAIL->PASS"
        if tag or n.get("pass") is not True:
            print(f"  {w[0]}: base={b.get('pass')} {b.get('score')}  new={n.get('pass')} {n.get('score')}{tag}")
    print(f"NEW pass {n_new}/{len(work)}   BASE pass {n_base}/{len(work)}", flush=True)

    if flips and flags:
        print("attribution (flag that restores PASS):")
        tasks = [((w[0], fl), exe, w[1], w[2], {**new_env, fl: "1"}) for w in flips for fl in flags]
        outs = run_many(tasks, jobs)
        for w in flips:
            for fl in flags:
                r = outs[(w[0], fl)]
                print(f"  {w[0]} with {fl}: pass={r.get('pass')} {r.get('score')}")

    if n_id:
        rest = [d for d in rows if d not in set(hits) and "error" not in rows[d]]
        random.seed(0)
        sample = random.sample(rest, min(n_id, len(rest)))
        with ThreadPoolExecutor(max_workers=jobs) as pool:
            pairs = list(pool.map(lambda d: (d, dump(exe, rows[d]["path"]) == dump(base, rows[d]["path"])), sample))
        bad = [d for d, same in pairs if not same]
        for d in bad:
            print(f"  identity {d}: DIFFERS <<<")
        print(f"identity sample: {len(sample) - len(bad)}/{len(sample)} byte-identical")


if __name__ == "__main__":
    main()
